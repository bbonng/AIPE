# 문서 데이터 JSON 스키마

확정된 설계 내용은 아래 구조 그대로 JSON으로 정리한다. `make_docx.py`가 이 구조를 읽어 DOCX를 생성한다.

```json
{
  "meta": {
    "year": "2026",
    "semester": "1",
    "school": "OO중학교",
    "grade": "2",
    "teacher": "교사 표기명",
    "sport": "종목명",
    "totalHours": 17,
    "evalPeriod": "4~6월"
  },
  "part1": {
    "evalPlans": [
      {
        "title": "수행평가 명칭 (내용요소 + 수행요소)",
        "method": "실기형",
        "period": "평가 시기",
        "standards": ["[9체02-17] 성취기준 전체 문구"],
        "taskFlow": ["1) 단계", "2) 단계", "3) 단계"],
        "elements": [{"name": "평가요소명", "score": 40}],
        "aiPolicy": "AI 활용 기준 문구",
        "rubric": [
          {
            "element": "평가요소명 (elements의 name과 정확히 일치)",
            "grades": [
              {"score": 40, "criteria": "최고 수준 채점기준"},
              {"score": 34, "criteria": "다음 수준"},
              {"score": 28, "criteria": "다음 수준"},
              {"score": 22, "criteria": "최저 수준"}
            ]
          }
        ]
      }
    ]
  },
  "part2": {
    "rows": [
      {
        "period": "4월 1주",
        "hours": 2,
        "unit": "단원명(주제)\n[핵심 아이디어]",
        "standard": "[9체02-17]",
        "evalType": "형성평가 또는 [수행평가1]",
        "detail": {
          "핵심개념": "…",
          "핵심탐구질문": "… (모든 차시에서 서로 달라야 함)",
          "개별화전략": "어려운 학생 → … / 숙달 학생 → …",
          "피드백전략": "…",
          "형성평가": "도입 발문 또는 마무리 발문",
          "수행지시어": "(수행평가 차시에만) 무엇을 어떻게 평가하는지"
        }
      }
    ]
  },
  "part3": {
    "semester": [
      {"level": "A", "desc": "…", "rate": "80% 이상"},
      {"level": "B", "desc": "…", "rate": "60% 이상 ~ 80% 미만"},
      {"level": "C", "desc": "…", "rate": "60% 미만"}
    ],
    "byStandard": [
      {"unit": "단원명", "standard": "[코드] 성취기준 전체 문구", "A": "…", "B": "…", "C": "…"}
    ]
  }
}
```

## 규칙

- `rubric[].element`는 `elements[].name`과 **문자열이 정확히 일치**해야 한다.
- 루브릭 `grades`의 점수는 **내림차순**, 첫 점수는 해당 요소의 배점과 같다.
- `part2.rows`의 `hours` 합계는 `meta.totalHours`와 같아야 한다.
- 선택한 모든 성취기준 코드는 `part2.rows`의 `standard`에 최소 1회 등장해야 한다.
- `핵심탐구질문`은 전 차시에 걸쳐 중복 금지.
- JSON만 출력할 때는 마크다운 코드블록 없이 순수 JSON으로 출력한다 (make_docx.py가 코드블록도 걸러내긴 함).
