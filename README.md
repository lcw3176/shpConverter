# shpConverter
좌표를 주소 혹은 주소를 좌표로 변환 후 shp 파일 생성

### 주소 -> 좌표 사용법
* 엑셀 A열에는 지도에 표시될 이름을, B열에는 상세 주소를 적음
* 변환할 좌표계를 설정 후 실행, 지원되는 좌표계 ('EPSG:4326', 'EPSG:3857', 'EPSG:5174', 'EPSG:5179')
* .xlsx, .csv, .html, .shp, .shx, .dbf 파일 생성됨.
* .xlsx, .csv: 좌표가 추가된 엑셀 파일
* .html: 생성된 좌표 기준의 지도 미리보기
* .shp, .shx, .dbf: GIS 프로그램을 위한 파일

### 좌표 -> 주소 사용법
* 엑셀 A열에는 지도에 표시될 이름, B열에는 경도, C열에는 위도를 적음
* 현재 자신에게 맞는 좌표계를 설정 후 변환 시작
* 생성된 엑셀 파일에는 도로명 주소와 지번 주소가 기존 좌표와 표기됨
* 생성되는 파일은 주소->좌표와 같음

