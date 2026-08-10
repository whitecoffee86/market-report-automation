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

    # 비교단지 라벨 5개 (고가/저가 차트 옆)
    label_ids = [79, 80, 81, 82, 83]
    for sid, item in zip(label_ids, d["stock_labels"]):
        set_shape_text(S(sid), f"{item['name']} [{item['built']}, {item['units']}세대]")

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
    """chart3.xml (stockChart, 고가/저가/평균가 x 5개 동) - python-pptx 미지원이라 XML 직접 수정."""
    tmp_dir = "unpacked_stock_tmp"
    shutil.rmtree(tmp_dir, ignore_errors=True)
    with zipfile.ZipFile(pptx_path) as z:
        z.extractall(tmp_dir)

    chart_path = f"{tmp_dir}/ppt/charts/chart3.xml"
    xml = open(chart_path, encoding="utf-8").read()

    dongs = d["stock_dongs"]  # list of 5 dong names
    series = d["stock_series"]  # dict: high, low, avg -> list of 5 values

    # 카테고리(동 이름) 치환: 5개 동 이름이 반복 등장(series 3개 x 5개 = 15회)
    old_dongs = re.findall(r"<c:pt idx=\"\d+\"><c:v>([^<]+)</c:v></c:pt>", xml)
    # 카테고리 strCache 블록만 골라 교체 (숫자가 아닌 값)
    def repl_cat_block(match):
        block = match.group(0)
        for i, name in enumerate(dongs):
            block = re.sub(
                rf'(<c:pt idx="{i}"><c:v>)[^<]*(</c:v></c:pt>)',
                rf"\g<1>{name}\g<2>",
                block,
                count=1,
            )
        return block

    xml = re.sub(r"<c:cat>.*?</c:cat>", repl_cat_block, xml, flags=re.S)

    # 값(numCache) 치환: 시리즈 순서는 고가/저가/평균가 (chart3 원본 구조 기준)
    ser_blocks = re.findall(r"<c:ser>.*?</c:ser>", xml, flags=re.S)
    order = ["high", "low", "avg"]
    new_xml = xml
    for ser_xml, key in zip(ser_blocks, order):
        vals = series[key]
        new_ser = ser_xml
        for i, v in enumerate(vals):
            new_ser = re.sub(
                rf'(<c:val>.*?<c:pt idx="{i}"><c:v>)[^<]*(</c:v>)',
                rf"\g<1>{v}\g<2>",
                new_ser,
                count=1,
                flags=re.S,
            )
        new_xml = new_xml.replace(ser_xml, new_ser, 1)

    with open(chart_path, "w", encoding="utf-8") as f:
        f.write(new_xml)

    out_tmp = pptx_path + ".tmp"
    if zipfile_exists := True:
        import os
        if os.path.exists(out_tmp):
            os.remove(out_tmp)
    with zipfile.ZipFile(out_tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in __import__("os").walk(tmp_dir):
            for file in files:
                full = __import__("os").path.join(root, file)
                rel = __import__("os").path.relpath(full, tmp_dir)
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
