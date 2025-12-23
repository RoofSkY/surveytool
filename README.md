# Survey Statistics Tool (만족도 조사 통계 프로그램)

이 프로그램은 설문 조사 응답 데이터를 입력받아 **빈도수**와 **응답 비율**을 엑셀 파일로 자동 생성해주는 파이썬 도구입니다. 데이터의 형태에 따라 **가로형(Horizontal)**과 **세로형(Vertical)** 두 가지 버전을 제공합니다.

## 빌드 및 설치 방법 (Build & Installation)

이 프로그램을 실행하기 위해서는 파이썬(Python 3.x)과 관련 라이브러리가 필요합니다.

### 1. 필수 라이브러리 설치

터미널(또는 CMD)에서 아래 명령어를 입력하여 필요한 패키지를 설치합니다.

```bash
pip install pandas xlsxwriter
```

### 2. 소스 코드 다운로드

저장소를 클론하거나 파이썬 파일(`surveytool_horizontal.py`, `surveytool_vertical.py`)을 다운로드합니다.

```bash
git clone https://github.com/RoofSkY/surveytool.git
cd ${폴더이름}
```

## 실행 파일(.exe) 빌드 방법

파이썬이 설치되지 않은 환경에서도 사용할 수 있도록 `PyInstaller`를 이용해 실행 파일로 변환하는 방법입니다.

### 1. PyInstaller 설치

터미널에서 아래 명령어를 입력하여 빌드 도구를 설치합니다.

```bash
pip install pyinstaller
```

### 2. EXE 파일 생성

각 스크립트를 단일 실행 파일로 빌드합니다. (터미널에서 해당 프로젝트 폴더로 이동 후 실행)

-   **가로형 빌드:**

```bash
pyinstaller --onefile surveytool_horizontal.py
```

-   **세로형 빌드:**

```bash
pyinstaller --onefile surveytool_vertical.py
```

### 3. 결과 확인

1. 빌드가 완료되면 폴더 내에 `dist`라는 디렉토리가 생성됩니다.
2. `dist` 폴더 안에 생성된 `surveytool_horizontal.exe` 또는 `surveytool_vertical.exe` 파일을 배포하여 사용하면 됩니다.

> **참고**: 빌드 시 `--noconsole` 옵션을 사용하지 않는 것을 권장합니다. 이 프로그램은 사용자로부터 입력을 받는 **콘솔 창(터미널)이 반드시 필요**하기 때문입니다.

## 기능 설명 (Features)

이 프로그램은 입력받은 데이터를 바탕으로 다음과 같은 통계 처리를 수행합니다:

-   **데이터 집계**: 각 문항별 선택지의 빈도수를 계산합니다.
-   **비율 계산**: (선택된 횟수 / 해당 문항의 총 선택 수)를 통해 백분율을 도출합니다.
-   **자동 서식 적용**:
-   **가로형**: Pandas DataFrame 구조를 활용해 엑셀 표 형태로 깔끔하게 저장합니다.
-   **세로형**: 문항별로 행을 나누고 배경색과 테두리 서식을 적용하여 가독성을 높입니다.

-   **중복 응답 처리**: 한 문항에 여러 번호를 입력(예: `1,3`)하더라도 각각의 빈도에 정확히 반영됩니다.

## 사용 방법 (Usage)

1. **설문 설정**: 프로그램 실행 후 총 문항 수와 각 문항의 선택지 개수를 입력합니다.
2. **응답 데이터 입력**:

-   문항 사이는 **띄어쓰기**로 구분합니다.
-   한 문항 내 중복 선택은 **쉼표(,)**로 구분합니다.
-   입력 예시: `1,2 3 5` (1번 문항 1,2번 중복 / 2번 문항 3번 / 3번 문항 5번 선택).

3. **완료 및 저장**: 모든 응답을 입력한 후 `q`를 입력하면 즉시 엑셀 파일이 생성됩니다.

## 파일 구조 설명

-   `surveytool_horizontal.py`: Pandas를 사용하여 데이터를 처리하며, 엑셀 시트에서 문항이 가로 방향(Column)으로 배치됩니다.
-   `surveytool_vertical.py`: xlsxwriter를 사용하여 직접 서식을 지정하며, 문항이 세로 방향(Row)으로 배치되어 한눈에 보기 편한 보고서 형태를 제공합니다.

## License

MIT License
