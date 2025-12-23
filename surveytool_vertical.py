import pandas as pd
import os
import xlsxwriter


def get_integer_input(prompt):
    while True:
        try:
            value = int(input(prompt))
            if value < 1:
                print("1 이상의 정수를 입력해주세요.")
                continue
            return value
        except ValueError:
            print("숫자만 입력해주세요.")


def main():
    print("=== 만족도 조사 통계 프로그램 (세로형) ===")

    try:
        num_questions = get_integer_input("총 문항 수는 몇 개입니까?: ")
    except KeyboardInterrupt:
        return

    question_configs = []
    for i in range(num_questions):
        max_choice = get_integer_input(f"{i+1}번 문항의 선택지 개수: ")
        question_configs.append(max_choice)

    print("\n=== 응답 데이터 입력 ===")
    print("1. 문항 간의 구분은 '공백(스페이스바)'으로 합니다.")
    print("2. 중복 응답(복수 선택)은 '쉼표(,)'로 구분하며 띄어쓰기 없이 입력합니다.")
    print("   (예시: 1번 문항(1,3번 선택), 2번 문항(5번 선택) -> 1,3 5)")
    print("'q'를 입력하면 입력을 종료하고 엑셀을 생성합니다.\n")

    responses = []
    person_count = 0

    while True:
        user_input = input(f"응답자 {person_count + 1} ('q' 종료) >> ")
        if user_input.lower() in ['q', 'exit']:
            break
        if not user_input.strip():
            continue

        try:
            raw_answers = user_input.split()
            if len(raw_answers) != num_questions:
                print(f"[오류] 문항 수 불일치")
                continue

            parsed_row = []
            valid_input = True
            for idx, ans_str in enumerate(raw_answers):
                selections = [int(x) for x in ans_str.split(',')]
                for sel in selections:
                    if sel < 1 or sel > question_configs[idx]:
                        valid_input = False
                        break
                if not valid_input:
                    break
                parsed_row.append(selections)

            if not valid_input:
                print("[오류] 범위를 벗어난 응답이 있습니다.")
                continue

            responses.append(parsed_row)
            person_count += 1
        except:
            print("[오류] 형식이 잘못되었습니다.")

    if person_count == 0:
        return

    output_filename = "survey_result_vertical.xlsx"

    try:
        workbook = xlsxwriter.Workbook(output_filename)

        # 스타일 설정
        header_fmt = workbook.add_format(
            {'bold': True, 'align': 'center', 'bg_color': '#F2F2F2', 'border': 1})
        percent_fmt = workbook.add_format(
            {'num_format': '0.0%', 'align': 'center', 'border': 1})
        count_fmt = workbook.add_format(
            {'num_format': '#,##0', 'align': 'center', 'border': 1})
        normal_fmt = workbook.add_format({'align': 'center', 'border': 1})

        # 두 개의 시트 생성
        sheets = [
            {'name': '빈도수_결과', 'type': 'count'},
            {'name': '응답비율_결과', 'type': 'percent'}
        ]

        for s_info in sheets:
            ws = workbook.add_worksheet(s_info['name'])
            curr_row = 0

            for i in range(num_questions):
                max_opt = question_configs[i]
                counts = [0] * max_opt

                # 집계
                for res in responses:
                    for choice in res[i]:
                        counts[choice-1] += 1

                total_selections = sum(counts)

                # 첫 번째 행: 문항 및 선택지 헤더
                ws.write(curr_row, 0, f"문항{i+1}", header_fmt)
                for j in range(max_opt):
                    ws.write(curr_row, j + 1, f"선택지{j+1}", header_fmt)

                # 두 번째 행: 실제 데이터 (빈도수 또는 백분율)
                data_label = "빈도수" if s_info['type'] == 'count' else "백분율"
                ws.write(curr_row + 1, 0, data_label, header_fmt)

                for j in range(max_opt):
                    if s_info['type'] == 'count':
                        ws.write(curr_row + 1, j + 1, counts[j], count_fmt)
                    else:
                        val = counts[j] / \
                            total_selections if total_selections > 0 else 0
                        ws.write(curr_row + 1, j + 1, val, percent_fmt)

                curr_row += 3  # 문항 간 간격 (데이터 행 다음 한 줄 띄움)

            ws.set_column(0, max(question_configs), 12)  # 열 너비 자동 조정

        workbook.close()
        print(f"\n[성공] '{output_filename}' 생성 완료!")
        print("- 시트1: 빈도수 결과 (가로형)")
        print("- 시트2: 응답비율 결과 (가로형)")

    except Exception as e:
        print(f"오류 발생: {e}")

    input("\n엔터를 누르면 종료됩니다.")


if __name__ == "__main__":
    main()
