# 시장조사 보고서 자동 생성 v2 (GitHub Actions)

260810_시장조사_양식_개선_v_2_0.pptx 기준. 사용법은 기존 v1과 동일합니다
(data/ 폴더에 json 올리면 자동으로 output/ 에 pptx 생성).

## v1과 다른 점 (데이터 스키마 변경사항)

- `region_name`, `region_avg_eok` 필드 **삭제됨** (v2 템플릿에 해당 요소 없음)
- `judge_table`는 이제 3개 키만 사용: `location`(입지및시장환경), `product`(상품), `brand`(브랜드영향)
  - 값은 "양호" / "보통" / "주의" 중 하나 (v1의 "신중"이 "주의"로 변경됨)
- `final_grade`, `final_grade_note` 필드 **추가됨** (예: "A", "D+4M" → "A 등급(D+4M)"으로 표시)
- `grade_table`은 v1과 동일한 5개 키(supply/price/deal/presale/unsold), 값은 S/A/B/C/D
  - S, A 등급은 파란 배지, 그 외는 연한 배지로 자동 표시
- `compare1_built`, `compare2_built` **추가됨** (선택, 비교단지 입주월 표시용)
- `opinions`: 검토의견을 **배열**로 입력 (문장 개수 자유, 2개든 5개든 가능)
  - 기존 `opinion_1/2/3` 방식도 호환되지만 새로 만들 땐 `opinions` 배열 사용 권장
- 나머지 필드(title, date, region_grade, 가격, 입지코멘트, stock_dongs/labels/series 등)는 v1과 동일

## 주의사항

- 문장이 너무 길면(특히 opinions 5줄 이상) 표/차트와 살짝 겹칠 수 있습니다.
  실제 PowerPoint에서 열어 텍스트박스 크기를 살짝 조정하시면 됩니다.
- 비교단지(막대그래프/오각형차트)는 템플릿 레이아웃상 **정확히 2개**로 고정되어 있습니다
  (고가/저가 비교의 동 개수처럼 가변으로 만들려면 템플릿 자체 슬롯 추가가 필요합니다).
