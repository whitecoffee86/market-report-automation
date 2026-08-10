#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
용인 신봉2지구 시장조사 보고서 템플릿 자동 채우기 스크립트

사용법:
    python3 generate_report.py data.json output.pptx

data.json 예시는 sample_data.json 참고.
"""
import os
import sys
import json
import re
import zipfile
import shutil
import copy
from pptx import Presentation
from pptx.chart.data import CategoryChartData

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.pptx")


def find_by_id(shapes, sid):
    for shp in shapes:
        if shp.shape_id == sid:
            return shp
        if shp.shape_type == 6:  # GROUP
            r = find_by_id(shp.shapes, sid)
            if r is not None:
                return r
    return None


def set_para_text(paragraph, new_text):
    runs = paragraph.runs
    if not runs:
        return
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""


def set_shape_text(shape, new_text, para_idx=0):
    set_para_text(shape.text_frame.paragraphs[para_idx], new_text)


def clear_shape_text(shape):
    for p in shape.text_frame.paragraphs:
        set_para_text(p, "")


def fill_text(slide, d):
    S = lambda sid: find_by_id(slide.shapes, sid)

    # 제목 / 작성일
    set_shape_text(S(3), f"{d['title']} 시장조사 보고서")
    set_shape_text(S(4), f"’{d['date']}  ㅣ  건축·주택마케팅팀")

    # 지역등급 뱃지 (그룹 68 내부 #70)
    set_shape_text(S(70), d["region_grade"])

    # 적정분양가
    set_shape_text(S(90), f"@{d['compare1_price_py']}")
    set_shape_text(S(92), f"@{d['compare2_price_py']}")
    set_shape_text(S(94), f"@{d['site_price_py']}")
    set_shape_text(S(96), f"{d['site_price_eok']}억  (@{d['site_price_py']}만원)")
    set_shape_text(S(98), f"84㎡ ({d['pyeong']}평)기준 (확장포함 · 이자후불제)")

    # 입지 코멘트 (교통/생활/개발) - 3개 문단
    box = S(67)
    set_para_text(box.text_frame.paragraphs[0], f"교통\t{d['loc_transport']}")
    set_para_text(box.text_frame.paragraphs[1], f"생활\t{d['loc_life']}")
    set_para_text(box.text_frame.paragraphs[2], f"개발\t{d['loc_dev']}")

    # 지역 위계 서술 (2문단)
    box = S(46)
    set_para_text(box.text_frame.paragraphs[0], d["region_summary_1"])
    set_para_text(box.text_frame.paragraphs[1], d["region_summary_2"])

    # 종합의견 (3문단)
    box = S(78)
    set_para_text(box.text_frame.paragraphs[0], d["opinion_1"])
    set_para_text(box.text_frame.paragraphs[1], d["opinion_2"])
    set_para_text(box.text_frame.paragraphs[2], d["opinion_3"])

    # 지역위계 하단 서술 + 수지구 평균
    set_shape_text(S(64), d["hierarchy_note"])
    set_shape_text(S(72), f"<{d['region_name']} 평균  {d['region_avg_eok']}억>")

    # 비교단지 라벨 (고가/저가 차트 옆, 최대 5개 슬롯 - 원본 템플릿 레이아웃 한계)
    label_ids = [79, 80, 81, 82, 83]
    labels = d["stock_labels"]
    for i, sid in enumerate(label_ids):
        if i < len(labels):
            item = labels[i]
            set_shape_text(S(sid), f"{item['name']} [{item['built']}, {item['units']}세대]")
        else:
            clear_shape_text(S(sid))

    # 표 1 (수급/가격/거래/분양/미분양 등급) - shape_id 5, row 2
    tbl = S(5).table
    grades = d["grade_table"]  # dict: supply, price, deal, presale, unsold
    order = ["supply", "price", "deal", "presale", "unsold"]
    for col, key in enumerate(order, start=1):
        tbl.cell(2, col).text_frame.paragraphs[0].runs[0].text = grades[key]

    # 표 2 (입지/수급/브랜드/상품 판정) - shape_id 48, row 1
    tbl = S(48).table
    judge = d["judge_table"]  # dict: location, supply, brand, product
    order2 = ["location", "supply", "brand", "product"]
    for col, key in enumerate(order2):
        tbl.cell(1, col).text_frame.paragraphs[0].runs[0].text = judge[key]

    # 작성자 가이드 문구(안내용 콜아웃) 비우기 - 최종 보고서에는 불필요
    for gid in (13, 14, 19):
        grp = S(gid)
        for shp in grp.shapes:
            if shp.has_text_frame and shp.text_frame.text.strip():
                clear_shape_text(shp)


def fill_charts(slide, d):
    S = lambda sid: find_by_id(slide.shapes, sid)

    # 막대그래프 (chart1, shape 11): 비교단지1 / 비교단지2 / SITE 적정가
    chart = S(11).chart
    cd = CategoryChartData()
    cd.categories = [d["compare1_name"], d["compare2_name"], "SITE 적정가\n(신규)"]
    cd.add_series("가격", (d["compare1_price_eok"], d["compare2_price_eok"], d["site_price_eok"]))
    chart.replace_data(cd)

    # 오각형(레이더) 차트 (chart2, shape 77): 입지/브랜드/단지/가격수준/연식 x 3세트
    chart = S(77).chart
    cd = CategoryChartData()
    cd.categories = ["입지(35)", "브랜드(15)", "단지(40)", "가격수준(10)", "연식(50)"]
    r = d["radar"]
    cd.add_series(d["compare1_name"], tuple(r["compare1"]))
    cd.add_series(d["compare2_name"], tuple(r["compare2"]))
    cd.add_series("SITE", tuple(r["site"]))
    chart.replace_data(cd)


def fill_stock_chart(pptx_path, d):
    """chart3.xml (stockChart, 고가/저가/평균가 x N개 동, N=3~5) - python-pptx 미지원이라 XML 직접 수정.
    카테고리 개수가 원본(5개)과 다를 수 있으므로 c:pt 요소 자체를 지우고 다시 생성한다."""
    from lxml import etree

    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    NS = {"c": C_NS}

    tmp_dir = "unpacked_stock_tmp"
    shutil.rmtree(tmp_dir, ignore_errors=True)
    with zipfile.ZipFile(pptx_path) as z:
        z.extractall(tmp_dir)

    chart_path = f"{tmp_dir}/ppt/charts/chart3.xml"

    dongs = d["stock_dongs"]           # N개 동 이름
    series = d["stock_series"]         # dict: high, low, avg -> N개 값
    n = len(dongs)
    if n < 3 or n > 5:
        raise ValueError("고가/평균/저가 비교는 3~5개 동만 지원합니다 (원본 템플릿 레이아웃 제약).")

    tree = etree.parse(chart_path)
    root = tree.getroot()
    sers = root.findall(".//c:ser", NS)
    order = ["high", "low", "avg"]  # 원본 chart3.xml의 시리즈 순서 (고가/저가/평균가)

    def rebuild_pts(cache_el, values):
        pt_count = cache_el.find("c:ptCount", NS)
        pt_count.set("val", str(len(values)))
        for pt in cache_el.findall("c:pt", NS):
            cache_el.remove(pt)
        for i, val in enumerate(values):
            pt = etree.SubElement(cache_el, f"{{{C_NS}}}pt")
            pt.set("idx", str(i))
            v_el = etree.SubElement(pt, f"{{{C_NS}}}v")
            v_el.text = str(val)

    for ser, key in zip(sers, order):
        cat_cache = ser.find(".//c:cat//c:strCache", NS)
        rebuild_pts(cat_cache, dongs)

        num_cache = ser.find(".//c:val//c:numCache", NS)
        rebuild_pts(num_cache, series[key])

    tree.write(chart_path, xml_declaration=True, encoding="UTF-8", standalone=True)

    out_tmp = pptx_path + ".tmp"
    if os.path.exists(out_tmp):
        os.remove(out_tmp)
    with zipfile.ZipFile(out_tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for root_dir, _, files in os.walk(tmp_dir):
            for file in files:
                full = os.path.join(root_dir, file)
                rel = os.path.relpath(full, tmp_dir)
                zf.write(full, rel)
    shutil.move(out_tmp, pptx_path)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    if len(sys.argv) != 3:
        print("사용법: python3 generate_report.py data.json output.pptx")
        sys.exit(1)
    data_path, out_path = sys.argv[1], sys.argv[2]
    d = json.load(open(data_path, encoding="utf-8"))

    prs = Presentation(TEMPLATE)
    slide = prs.slides[0]
    fill_text(slide, d)
    fill_charts(slide, d)
    prs.save(out_path)

    fill_stock_chart(out_path, d)
    print(f"완료: {out_path}")


if __name__ == "__main__":
    main()
