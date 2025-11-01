import xlwings as xw
from datetime import datetime


LIGHT_GREEN = (226, 239, 218)


def _to_int_term(raw_value, issue_age, limit_age: int = 100) -> int:
    # '보험기간' 셀 값을 정수 기간으로 변환
    # - 값이 "종신"이면: limit_age - issue_age 계산
    # - 그 외에는 숫자로 변환(문자 "20", "20.0" 등 허용)
    text = str(raw_value).strip() if raw_value is not None else ""
    if not text:
        raise ValueError("보험기간 값이 비어 있습니다.")
    if text.lower() == "종신":
        if issue_age is None:
            raise ValueError("종신 계산을 위해 'C4'에 연령이 필요합니다.")
        return int(limit_age - int(float(issue_age)))
    return int(float(text))

def sync_layout() -> None:
    # '입력'!C5(보험기간)을 기준으로 '결과출력' 시트 레이아웃 갱신
    # xlwings를 통해 VBA에서 호출:
    #   RunPython "import mymodule; mymodule.sync_layout()"
    wb = xw.Book.caller()
    app = wb.app

    inp = wb.sheets["입력"]
    out = wb.sheets["결과출력"]

    # 디버그 마커: 파이썬 호출 시점과 C5 값을 눈으로 확인
    out["Z1"].value = f"PY {datetime.now():%Y-%m-%d %H:%M:%S}"
    out["Z2"].value = f"C5={inp['C5'].value}"

    # 읽기 전에 엑셀 계산 강제 갱신
    try:
        app.calculate()
    except Exception:
        pass

    issue_age = inp["C4"].value
    raw_term = inp["C5"].value
    term = _to_int_term(raw_term, issue_age)
    if term <= 0:
        raise ValueError("보험기간은 1 이상이어야 합니다.")

    # A열 라벨 목록
    labels = [
        "구분",
        "PV(유입)_보험료",
        "PV(유출)_보험금",
        "최선추정부채(BEL)",
        "위험조정(RA)",
        "이행현금흐름(FCF)",
        "보험계약마진(CSM)",
        "보험계약부채",
    ]

    num_rows = len(labels)
    num_cols = 1 + term  # A열 포함 총 열 수

    # 라벨/헤더 쓰기(정확 범위만)
    out["A1"].resize(num_rows, 1).value = [[label] for label in labels]
    out["B1"].resize(1, term).value = [[f"{i}차년도" for i in range(1, term + 1)]]

    # 우측 잔여 헤더 셀 정리(이전 차년도 텍스트가 남지 않도록)
    _EXTRA = 200  # 현재 기간 오른쪽으로 비울 최대 열 수
    out["B1"].offset(0, term).resize(1, _EXTRA).value = [[""] * _EXTRA]
    out["B1"].offset(0, term).resize(1, _EXTRA).color = None

    # 배경색(라벨 열, 헤더 행)
    out["A1"].resize(num_rows, 1).color = LIGHT_GREEN
    out["B1"].resize(1, term).color = LIGHT_GREEN

    # 표 범위에만 얇은 테두리 적용
    table = out["A1"].resize(num_rows, num_cols)
    xlContinuous, xlThin = 1, 2
    xlLineStyleNone = -4142

    for idx in (7, 8, 9, 10):  # 좌, 상, 하, 우
        b = table.api.Borders(idx)
        b.LineStyle = xlContinuous
        b.Weight = xlThin

    if num_cols > 1:
        b = table.api.Borders(11)  # 내부 세로선
        b.LineStyle = xlContinuous
        b.Weight = xlThin

    if num_rows > 1:
        b = table.api.Borders(12)  # 내부 가로선
        b.LineStyle = xlContinuous
        b.Weight = xlThin

    # 포함 영역 전체에 '모든 테두리' 적용(엑셀의 All Borders와 동일)
    try:
        table.api.Borders.LineStyle = xlContinuous
        table.api.Borders.Weight = xlThin
    except Exception:
        pass

    # 우측 잔여 영역의 테두리를 제거(이전 테두리가 남지 않도록)
    _EXTRA = 200
    residual = out["A1"].offset(0, num_cols).resize(num_rows, _EXTRA)
    # 중요: 잔여 영역의 왼쪽 테두리(7)는 표의 오른쪽 테두리와 공유되므로 지우지 않음
    for idx in (8, 9, 10, 11, 12):
        try:
            residual.api.Borders(idx).LineStyle = xlLineStyleNone
        except Exception:
            pass

    # 표 외곽 테두리를 다시 한 번 적용해 오른쪽 경계가 확실히 보이도록 함
    for idx in (7, 8, 9, 10):
        b = table.api.Borders(idx)
        b.LineStyle = xlContinuous
        b.Weight = xlThin

    # 선택: 최종 숫자 기간을 C5에 반영(초기 입력이 '종신'이었던 경우 유용)
    inp["C5"].value = term


