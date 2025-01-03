# cmdb-xlsx2csv

## Feature

- 다운로드 받은 구성관리조회 엑셀파일의 제목인 1행부터 2행까지 삭제 (A열부터 CP열까지 병합된 상태) 기능
- encoding 선택 기능 (euc-kr, utf-8)
- drag & drop 하면 자동 변환
- 데이터 정제 기능
  - None, NaN 처리
  - 모든 데이터를 문자열로 변환
  - non-breaking space를 일반 공백으로 변환
  - 모든 종류의 공백문자를 일반 공백으로 변환
  - 셀 내 줄바꿈 (Alt+Enter) 제거
  - 특수문자 제거 또는 변환

## How to Build

- upx https://github.com/upx/upx/releases

```sh
$ pip install pyinstaller
$ pyinstaller -F -w --add-data "logo.png;." cmdb-xlsx2csv.py
$ pyinstaller --upx-dir c:\upx424 -F -w --add-data "logo.png;." cmdb-xlsx2csv.py
```