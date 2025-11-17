#!/usr/bin/env python3
"""
이화방어기제검사 자동화 스크립트

Google Sheets 응답 → 채점 엑셀 자동 입력

Requirements:
    pip install gspread oauth2client openpyxl

사용 방법:
    python3 edmt_automation.py --sheet-url "YOUR_GOOGLE_SHEETS_URL" --excel-template "채점표.xlsx" --output "결과.xlsx"
"""

import argparse
import sys
from pathlib import Path

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    from openpyxl import load_workbook
except ImportError as e:
    print(f"❌ 필수 라이브러리가 설치되지 않았습니다: {e}")
    print("\n다음 명령어로 설치하세요:")
    print("pip install gspread oauth2client openpyxl")
    sys.exit(1)


# 20개 방어기제
DEFENSE_MECHANISMS = [
    "허세", "반동형성", "동일시", "수동공격", "투사", "전치", "부정", "통제", "억제", "왜곡",
    "예견", "합리화", "해리", "신체화", "승화", "행동화", "이타주의", "퇴행", "유머", "회피"
]


class EDMTAutomation:
    def __init__(self, credentials_file="credentials.json"):
        """
        초기화

        Args:
            credentials_file: Google API 인증 파일 경로
        """
        self.credentials_file = credentials_file
        self.client = None

    def authenticate(self):
        """Google Sheets API 인증"""
        try:
            scope = [
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive'
            ]
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_file, scope
            )
            self.client = gspread.authorize(creds)
            print("✅ Google Sheets API 인증 성공")
            return True
        except FileNotFoundError:
            print(f"❌ 인증 파일을 찾을 수 없습니다: {self.credentials_file}")
            print("\nGoogle Cloud Console에서 서비스 계정 키를 생성하고")
            print("credentials.json 파일로 저장하세요.")
            return False
        except Exception as e:
            print(f"❌ 인증 실패: {e}")
            return False

    def read_responses(self, sheet_url):
        """
        Google Sheets에서 응답 읽기

        Args:
            sheet_url: Google Sheets URL

        Returns:
            list: 응답 데이터 리스트
        """
        try:
            # Sheets 열기
            sheet = self.client.open_by_url(sheet_url)
            worksheet = sheet.get_worksheet(0)  # 첫 번째 시트

            # 모든 데이터 가져오기
            all_records = worksheet.get_all_records()

            print(f"✅ {len(all_records)}개의 응답을 읽어왔습니다")
            return all_records

        except Exception as e:
            print(f"❌ Google Sheets 읽기 실패: {e}")
            return []

    def parse_response(self, response_row):
        """
        응답 데이터 파싱

        Args:
            response_row: 한 행의 응답 데이터 (dict)

        Returns:
            dict: 파싱된 데이터
        """
        # 기본 정보 추출
        name = response_row.get('이름', '')
        age = int(response_row.get('나이', 0))
        gender = response_row.get('성별', '')

        # 200개 문항 응답 추출 (① ②  ③ ④ ⑤ → 1, 2, 3, 4, 5)
        answers = []
        for i in range(1, 201):
            question_key = f'{i}. '  # "1. 일생에 한 번뿐인..." 형식
            # 실제 키를 찾기
            full_key = None
            for key in response_row.keys():
                if key.startswith(question_key):
                    full_key = key
                    break

            if full_key:
                answer_text = response_row[full_key]
                # "① 전혀 아니다" → 1
                if '①' in answer_text:
                    score = 1
                elif '②' in answer_text:
                    score = 2
                elif '③' in answer_text:
                    score = 3
                elif '④' in answer_text:
                    score = 4
                elif '⑤' in answer_text:
                    score = 5
                else:
                    score = 0  # 무응답
                answers.append(score)
            else:
                answers.append(0)  # 키를 찾지 못한 경우

        return {
            'name': name,
            'age': age,
            'gender': gender,
            'answers': answers
        }

    def calculate_raw_scores(self, answers):
        """
        원점수 계산

        Args:
            answers: 200개 문항 응답 (list)

        Returns:
            dict: 각 방어기제별 원점수
        """
        scores = {}

        for mech_idx, mechanism in enumerate(DEFENSE_MECHANISMS):
            # 해당 방어기제의 문항 번호들 (1-based)
            item_numbers = [mech_idx + 1 + (i * 20) for i in range(10)]

            # 원점수 합산
            total_score = sum(answers[item_num - 1] for item_num in item_numbers)
            scores[mechanism] = total_score

        return scores

    def get_norm_table_code(self, age, gender):
        """
        나이와 성별에 따른 규준표 코드 결정

        Args:
            age: 나이
            gender: 성별 ("남" or "여")

        Returns:
            str: 규준표 이름 (예: "규준표11")
        """
        # 나이 그룹 결정
        if 15 <= age <= 29:
            age_code = 1
        elif 30 <= age <= 39:
            age_code = 2
        elif 40 <= age <= 49:
            age_code = 3
        else:  # 50세 이상
            age_code = 4

        # 성별 코드
        gender_code = 1 if gender == "남" else 2

        return f"규준표{age_code}{gender_code}"

    def convert_to_sten(self, raw_score, norm_table_data, mechanism_col):
        """
        원점수를 스텐점수로 변환

        Args:
            raw_score: 원점수
            norm_table_data: 규준표 데이터
            mechanism_col: 방어기제 열 인덱스

        Returns:
            int: 스텐점수 (1-10)
        """
        # 규준표에서 스텐점수 찾기
        # Row 2-11이 스텐점수 1-10에 해당
        for sten_score in range(1, 11):
            row_idx = sten_score + 1  # 엑셀 행 번호 (2-11)
            threshold = norm_table_data[row_idx][mechanism_col]

            if raw_score <= threshold:
                return sten_score

        return 10  # 최대값

    def write_to_excel(self, template_path, output_path, parsed_data, raw_scores):
        """
        채점 엑셀에 데이터 입력

        Args:
            template_path: 템플릿 엑셀 경로
            output_path: 출력 엑셀 경로
            parsed_data: 파싱된 응답 데이터
            raw_scores: 원점수 딕셔너리
        """
        try:
            # 엑셀 파일 로드
            wb = load_workbook(template_path)

            # 응답지_1 시트에 응답 입력
            response_sheet = wb['응답지_1']

            # 기본 정보 입력 (행 4)
            response_sheet['D4'] = parsed_data['name']  # 이름
            response_sheet['F4'] = parsed_data['age']    # 나이
            response_sheet['H4'] = parsed_data['gender'] # 성별

            # 200개 응답 입력
            answers = parsed_data['answers']
            for i, answer in enumerate(answers):
                # 20문항씩 나눠서 입력 (행 7, 10, 13, 16, 19, 22, 25, 28, 31, 34)
                section = i // 20
                position_in_section = i % 20

                # 행 번호 계산
                base_row = 7 + (section * 3)
                # 열 번호 계산 (D열부터 시작, 20개)
                col = 4 + position_in_section  # D=4

                response_sheet.cell(row=base_row, column=col, value=answer)

            # 저장
            wb.save(output_path)
            print(f"✅ 엑셀 파일 저장 완료: {output_path}")

        except Exception as e:
            print(f"❌ 엑셀 저장 실패: {e}")

    def process_all_responses(self, sheet_url, template_path, output_dir="./results"):
        """
        모든 응답 자동 처리

        Args:
            sheet_url: Google Sheets URL
            template_path: 템플릿 엑셀 경로
            output_dir: 출력 디렉토리
        """
        # 인증
        if not self.authenticate():
            return

        # 응답 읽기
        responses = self.read_responses(sheet_url)

        if not responses:
            print("❌ 처리할 응답이 없습니다")
            return

        # 출력 디렉토리 생성
        Path(output_dir).mkdir(exist_ok=True)

        # 각 응답 처리
        for idx, response_row in enumerate(responses, 1):
            print(f"\n{'=' * 60}")
            print(f"응답 {idx}/{len(responses)} 처리 중...")
            print(f"{'=' * 60}")

            # 응답 파싱
            parsed = self.parse_response(response_row)
            print(f"📝 이름: {parsed['name']}, 나이: {parsed['age']}, 성별: {parsed['gender']}")

            # 원점수 계산
            raw_scores = self.calculate_raw_scores(parsed['answers'])
            print("📊 원점수 계산 완료")

            # 엑셀에 저장
            output_filename = f"{parsed['name']}_{parsed['age']}세_{parsed['gender']}_결과.xlsx"
            output_path = Path(output_dir) / output_filename

            self.write_to_excel(template_path, str(output_path), parsed, raw_scores)

        print(f"\n{'=' * 60}")
        print(f"✅ 전체 {len(responses)}개 응답 처리 완료!")
        print(f"📁 결과 파일 위치: {output_dir}/")
        print(f"{'=' * 60}")


def main():
    parser = argparse.ArgumentParser(description='이화방어기제검사 자동화')
    parser.add_argument('--sheet-url', required=True, help='Google Sheets URL')
    parser.add_argument('--template', required=True, help='채점 템플릿 엑셀 파일 경로')
    parser.add_argument('--output-dir', default='./results', help='출력 디렉토리')
    parser.add_argument('--credentials', default='credentials.json', help='Google API 인증 파일')

    args = parser.parse_args()

    # 자동화 실행
    automation = EDMTAutomation(credentials_file=args.credentials)
    automation.process_all_responses(
        sheet_url=args.sheet_url,
        template_path=args.template,
        output_dir=args.output_dir
    )


if __name__ == '__main__':
    main()
