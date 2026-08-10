# 시장조사 보고서 자동 생성 (GitHub Actions)

Claude 없이도, JSON 데이터만 올리면 자동으로 완성된 PPT가 나오는 구조입니다.

## 최초 1회 설정

1. GitHub에 새 저장소 생성 (예: `market-report-automation`), 이 폴더 전체를 push
2. 저장소 `Settings > Actions > General > Workflow permissions`에서
   **"Read and write permissions"** 선택 (자동 커밋을 위해 필요)

## 평소 사용법 (팀원 포함, 코드 지식 불필요)

1. `시장조사_보고서_입력폼.html`을 열어 데이터 입력
2. **"JSON 파일 다운로드"** 버튼 클릭 → `xxx.json` 파일 저장
3. GitHub 저장소의 `data/` 폴더로 이동 → **Add file → Upload files** → 방금 받은 json 업로드 → Commit
4. 자동으로 GitHub Actions가 실행되어 (1~2분 소요) `output/` 폴더에 완성된 `.pptx`가 생성됩니다
5. `output/` 폴더에서 파일을 열거나 다운로드

## 폴더 구조

```
market-report-automation/
├── template.pptx           # 원본 양식 (건드리지 마세요)
├── generate_report.py       # 데이터 1건 → PPT 1개 변환 로직
├── build_all.py              # data/ 안의 모든 json을 일괄 변환
├── requirements.txt
├── data/                     # 여기에 json을 올리면 자동 실행됨
└── output/                   # 완성된 PPT가 여기 쌓임
└── .github/workflows/generate.yml   # 자동화 설정
```

## 참고

- `data/` 폴더의 json 파일명이 그대로 결과 PPT 파일명이 됩니다 (예: `pangyo_2jigu.json` → `pangyo_2jigu.pptx`)
- 여러 명이 동시에 여러 json을 올려도 각각 별도 파일로 처리됩니다
- 템플릿 디자인을 바꾸고 싶으면 `template.pptx`를 교체 + `generate_report.py`의 shape_id 매핑을 다시 확인해야 합니다 (이 부분은 구조가 바뀌면 다시 문의해주세요)
