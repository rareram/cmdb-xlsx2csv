# cmdb-xlsx2csv

## UI

![cmdb-xlsx2csv-1](https://github.com/user-attachments/assets/c57ed28a-70b1-4235-aed6-2e57996a70ff)
![cmdb-xlsx2csv-2](https://github.com/user-attachments/assets/276d2094-f3f0-4295-b9fd-730fabd7d8cb)

## Feature

- 다운로드 받은 구성관리조회 엑셀파일의 제목인 1행부터 2행까지 삭제 (A열부터 CP열까지 병합된 상태) 기능
- encoding 선택 기능 (euc-kr, utf-8)
- drag & drop 하면 자동 변환
- 데이터 정제 기능
  - None, NaN 처리
  - 모든 데이터를 문자열로 변환
  - non-breaking space를 일반 공백으로 변환
  - 모든 종류의 공백문자를 일반 공백으로 변환
  - 셀 내 줄바꿈 (Alt+Enter) 제기
  - 특수문자 제거 또는 변환

## How to Build

```sh
pip install PySide6 pandas
```

- 윈도우에서 단일실행 파일 만들기
- upx https://github.com/upx/upx/releases

```sh
pip install pyinstaller
pyinstaller -F -w --add-data "logo.png;." cmdb-xlsx2csv.py
pyinstaller --upx-dir c:\upx424 -F -w --add-data "logo.png;." cmdb-xlsx2csv.py
```

## TroubleShooting

- 윈도우가 아닌 리눅스에서 PySide6로 실행시켰는데 에러가 나는 경우

```sh
# Debian/Ubuntu 기반
sudo apt-get install libxcb-cursor0

# RHEL/CentOS/RockyOS/Fedora 기반
sudo dnf install xcb-util-cursor
```