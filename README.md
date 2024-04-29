## Data Insight

### Data Insight.xlam 
  엑셀에서의 작업을 위한 매크로 모음. xlam 파일에 모아 둔 모듈, 사용자 폼, 리본 메뉴 양식
모듈 구분
1) 'Extract_' 타입 모듈들
2) 'Merge_' 타입 모듈들
3) 'RangeOps' 타입 모듈들
4) 'TextEdit' 타입 모듈들
5) 'newFucntion' 타입 모듈들


#### Extract_ 타입 모듈
  동일한 엑셀 템플릿 안에서 원하는 값들을 추출함

##### Extract_attributes_xlFofrmDs 
##### Extract_attributes_xlFormDs_sub_m
: 엑셀 이름 관리자를 통해 이름이 정의된 데이터 시트 양식에서 값들을 추출함

#### Merge_ 타입 모듈
  테이블 병합, 조인 등을 위한 모듈
  
##### Merge_byPriority
  우선순위 병합 : 두 개의 시트를 key와 column을 가지고 매핑하여 하나로 합치되, 겹치는 값에 대해서는 입력 우선순위를 정함(한쪽이 공란인 경우는 반대쪽을 취함)
##### Merge_EasyLookup
  데이터 업데이트(별도 read me 파일 참조)

#### RangeOps_ 타입모듈
  range 개체 조작을 위한 모듈
  
##### RangeOps_CompareValues
  두 개의 시트를 key column을 가지고 매핑하여 하나로 합치되, 겹치는 값에 대헛는 입력 우선위를 정함
##### RangeOps_InplaceRecursor
  셀 값 계산에서 재귀 호출을 1회 허용함
##### RangeOps_SwapRange
  두 개의 선택한 영역 사이의 값을 바꾼다
##### RangeOps_WhatsDifferent
  두 개의 테이블 사이에서 우선순위에 따라 업데이트


#### TextEidt_ 타입

##### TextEdit_AppendFront 
  셀 안의 텍스트 앞에 텍스트를 연결함
##### TextEdit_AppendBack 
  셀 안의 텍스트 뒤에 텍스트를 연결함
##### TextEdit_SliceFront(미구현) 
  특정한 문자(열)가 최초로 등장하는 지점까지 텍스트 자르기
##### TextEdit_SliceBack(미구현)
  특정한 문자(열)가 최초로 등장한 다음부터 텍스트 자르기
##### UseXlsxFunction(미구현 -> C# VSTO 검토) 
  : 엑셀의 기존 함수를 이용하는 방법
    - LEFT(C, 2)
    - RIGHT(C, 2)
    - MID(C, 2, 3)
    - TEXT(C, "")
    - SUBSTITUTE(C, "-", "")

#### newFucntion 타입

##### TEXTJOIN(미구현)
  excel 2016이하 버전에서 사용할 수 있는 TEXTJOIN 함수를 사용자 정의 함수로 구현
##### RCOUNTA
  빈셀이 ""일 경우 COUNTA가 제대로 작동하지 않는 문제를 해결함

### Data Insight - pyForData 
  효율적인 사무 처리를 위한 파이선 코드 조각. pandas를 이용하여 데이터 처리 시 자주 사용하는 기능을 클래스 또는 함수로 구현함

#### Monet.py
  Tesseract 엔진으로 동작하는 캡처&ocr 프로그램
#### ShapeShifter.py
  각종 포멧들의 변환(tif->pdf, pdf->tif, dxf->pdf, xlsx->pdf 등)
#### Table_cls.py
  테이블 클래스. 각종 테이블 형태 데이터에 대한 조작
#### common_queries.sql
  sqlite를 위한 쿼리 모음
#### process_batch_files.py 
  다량의 파일들에 대한 일괄 삭제, 복사, 이동 등
#### process_file_name.py 
  파일 이름 조작을 위한 간단한 파이썬 코드 조각
#### process_pdf_files.py 
  pdf 조작을 위한 파이썬 코드

### Data Insight - pyForDrawing

#### process_dxf_files.py
  dxf파일 조작을 위한 파이썬 코드

### Data Insight - Draft Master
.dwg, .dxf 파일에서 텍스트 정보를 중심으로 각종 데이터를 추출하기 위한 솔루션 패키지
### Data Insight - CollabControl 
공통기준문서나 공동작업문서의 버전관리, 변경점 공유, 백업, 로그 관리, 협업 등을 편리하게
### Data Insight - GuideBook 
분류체계, 속성체계 등에 대한 기본 교육, 실제 사례 학습. 각종 설비(고정, 회전, 배관, 계기, 전기) 분류 등에 대한 기본 소양교육. 당사의 솔루션과 상용 제품들의 활용에 대한 교육
### Data Insight - Freeware List 
데이터 정비 업무의 각 분야 및 문제에 적용할 수 있는 무료 소프트웨어들의 목록과 설명
### Data Insight - Support Process Management
MDM DBMS와 유사 또는 동일한 구조를 가지는 DBMS 시스템, 협업 및 태그의 추적 관리를 위한 데이터베이스, 작업용 폼 배포, 업로드, 정합성 검사
### Data Insight - AI
머신러닝 또는 딥러닝 기법을 응용한 솔루션

