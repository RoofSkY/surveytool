# Survey Statistics Tool (만족도 조사 통계 프로그램)

이 프로그램은 설문 조사 응답 데이터를 입력받아 **빈도수**와 **응답 비율**을 엑셀 파일로 자동 생성해주는 파이썬 도구입니다.

데이터의 형태에 따라 **가로형(Horizontal)**과 **세로형(Vertical)** 두 가지 버전을 제공합니다.

## Build & Installation

You need Python (Python 3.x) and related libraries to run this program.

### 1. Install pip

```bash
pip install pandas xlsxwriter
```

### 2. Install PyInstaller package

```bash
pip install pyinstaller
```

### 3. Create EXE files

-   **Horizontal Build:**

```bash
pyinstaller --onefile surveytool_horizontal.py
```

-   **Vertical Build:**

```bash
pyinstaller --onefile surveytool_vertical.py
```

### 4. Run the project

Distribute and use the file 'survetool_horizontal.exe' or 'survetool_vertical.exe' created in the 'dist' folder.

> **Note**: I recommend that you do not use the '--noconsole' option in the build, as this program must have a **console window (terminal) that receives input from the user**.

## Key Features

-   **데이터 집계**: 각 문항별 선택지의 빈도수를 계산합니다.
-   **비율 계산**: (선택된 횟수 / 해당 문항의 총 선택 수)를 통해 백분율을 도출합니다.
-   **자동 서식 적용**:
-   **가로형**: Pandas DataFrame 구조를 활용해 엑셀 표 형태로 깔끔하게 저장합니다.
-   **세로형**: 문항별로 행을 나누고 배경색과 테두리 서식을 적용하여 가독성을 높입니다.

-   **중복 응답 처리**: 한 문항에 여러 번호를 입력(예: `1,3`)하더라도 각각의 빈도에 정확히 반영됩니다.

## Example Usage

1. **설문 설정**: 프로그램 실행 후 총 문항 수와 각 문항의 선택지 개수를 입력합니다.
2. **응답 데이터 입력**:

-   문항 사이는 **띄어쓰기**로 구분합니다.
-   한 문항 내 중복 선택은 **쉼표(,)**로 구분합니다.
-   입력 예시: `1,2 3 5` (1번 문항 1,2번 중복 / 2번 문항 3번 / 3번 문항 5번 선택).

3. **완료 및 저장**: 모든 응답을 입력한 후 `q`를 입력하면 즉시 엑셀 파일이 생성됩니다.

## License

MIT License
