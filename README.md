# kapecp_SavePDF
본 코드는 [축산물원패스](https://www.ekape.or.kr/kapecp/ui/kapecp/index.html)에서 발급하는 한우의 등급판정서를 PDF로 자동 저장하는 프로그램입니다.
<br><br>
## 프로그램 실행 예시
![bandicam-2025-03-22-23-41-17-139](https://github.com/user-attachments/assets/e12d50b7-18f6-427a-9254-9130c8ddadcf)
<br><br>
## 결과
![그림1](https://github.com/user-attachments/assets/f7edc95c-1ee0-4b36-91a4-b1091d9d6cf8)


## 실행
```
python main.py
```
## 종속성 해결
이 requirements.txt파일은 본 코드가 의존하는 Python 라이브러리가 나열되어 있으며, 다음 같이 설치합니다.
```
pip install -r requirements.txt
```
## 개선할 사항
### - 페이지 로딩 중 클릭이 진행되었을 때, 에러로 인한 프로그램 종료
  * __이력번호 검색시 로딩이 2개가 존재함. 1차 로딩은 웹페이지 로딩이며, 두번째 로딩은 웹페이지 로딩이 끝난 후 이력번호의 각 정보를 Get 요청으로 받아올때까지 나타나는 로딩__,
  * __이에 웹페이지 로딩이 끝난 후 클릭이 진행될 수 있도록 개발했었지만, 두번째 로딩 요소의 태그명이 지속적으로 바뀌는 문제를 발견함__,
  * __결과적으로 동적으로 웹페이지가 모두 로딩되면 클릭 -> 약 7~8초의 대기 후 클릭하는 방식으로 수정함__,
  * __향후 최적화가 가능할 것으로 판단됨__,
