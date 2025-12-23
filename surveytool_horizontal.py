import pandas as pd
import os
# import xlsxwriter


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
    print("=== 만족도 조사 통계 프로그램 (가로형) ===")

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

    output_filename = "survey_result_horizontal.xlsx"

    # 1. 빈도수(건수) 계산
    max_option_overall = max(question_configs)
    stats_data = {f"문항 {i+1}": [0] *
                  max_option_overall for i in range(num_questions)}
    index_labels = [f"선택지 {i+1}" for i in range(max_option_overall)]

    for person_res in responses:
        for q_idx, selections in enumerate(person_res):
            for choice in selections:
                stats_data[f"문항 {q_idx+1}"][choice-1] += 1

    df_counts = pd.DataFrame(stats_data, index=index_labels)

    # 2. 백분율 계산: 각 문항(열)의 합계로 나눔
    # df_counts.sum(axis=0)은 각 문항별 총 선택 횟수입니다.
    df_percent = df_counts.div(df_counts.sum(axis=0), axis=1)
    df_percent = df_percent.fillna(0)  # 선택이 하나도 없는 문항의 경우 0으로 채움

    try:
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            # 빈도수 시트 저장
            df_counts.to_excel(writer, sheet_name='빈도수_결과')

            # 비율 시트 저장
            df_percent.to_excel(writer, sheet_name='응답비율_결과')

            workbook = writer.book
            worksheet_pct = writer.sheets['응답비율_결과']

            # 엑셀 퍼센트 서식 적용
            percent_fmt = workbook.add_format({'num_format': '0.0%'})
            worksheet_pct.set_column(1, num_questions, 12, percent_fmt)

        print(f"\n[성공] '{output_filename}' 생성 완료!")
        print("- 시트1: 빈도수 결과 (가로형)")
        print("- 시트2: 응답비율 결과 (가로형)")

    except Exception as e:
        print(f"오류 발생: {e}")

    input("\n엔터를 누르면 종료됩니다.")


if __name__ == "__main__":
    main()
