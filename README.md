# google_sheet_using_database

# Lib :
- gspread
- pyqt5
- numpy

--------------------

google sheet를 Database로 활용한 코드 입니다.

(This code utilizes google sheet as database.)

pyqt를 이용해서 GUI를 만들고 구글 시트와 연동시켰습니다.
(We created the GUI using pyqt and then we linked it to the Google sheet.)
--------------------


사용 환경 (Usage Environment): 
- 하루에 50 ~ 100건의 주문. 각 주문마다 약 30종류의 데이터. 
  (50 to 100 orders a day. About 30 kinds of data for each order.)

- 평균적으로 7-10명 동시 사용. (약 1년간 문제없이 사용)
  (On average, 7-10 people are used simultaneously)
  
-------------------------
사용법 (Usage)
  
- 주문을 등록하고, 고객들의 요구 및 회사 내부 사정에 따라 주문 데이터를 지속적으로 수정.
  (Register your order and continue to amend your order data based on customer needs and company internal conditions.)

- 등록한 시간을 10^(-6)초 까지 측정하여 key 값으로 사용
  (Measure the registered time up to 10^ (-6) seconds and use it as the key value.)


--------------------


클라우드 서비스 대신 구글 시트를 데이터베이스로 사용했을 때의 장점

1. 무제한 무료. 
2. 클라우드 서비스 못지 않은 서버의 안정성
3. 코드 짜기가 쉬움
4. 엑셀의 유용함을 그대로 이용할 수 있음.
5. 다수의 동시 작업 가능
6. 24시간 사용 가능.
7. 구글 시트는 항상 데이터를 백업 시켜준다.


--------------------

단점:

1. 클라우드 서비스보다 데이터의 안전성이 떨어질 수 있음.
- ex 1) 데이터 index 조회 후 해당 index의 data를 수정하는 경우. 조회와 수정 사이에 다른 요청이 들어와 data의 index를 변동 시킬 수 있음. 
      -> 아주 작은 확률이기도 하고 구글 시트는 자동 백업 기능을 지원하기 때문에 소규모 영업에 이용하기에는 무리없음.

 
 확실한 안전성을 위해서는 AWS나 google cloud에서 database를 사용하는 것이 좋습니다. 
 다만 스타트업이나, 영업량이 많지 않은 회사에서는 google sheet를 DB로 쓰는 것도 좋은 방법입니다.
 
 
     
