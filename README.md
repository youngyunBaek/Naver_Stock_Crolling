# Naver_Stock_Crolling
Self made for naver stock data crolling

 1. main.py
    - 5분마다 수집 작업을 수행하는 main 함수
    
    Line 8: 시간이 변경되었을때만 정보를 수집하기 위해 만든 변수 
            증권 시장을 기준으로 동작하다 보니 5분마다 동작 할 경우 필요없는 값이 수집되는 경우가 있어 증권 개장 시간에만 동작
            (비교가 되는 return 값에 경우 timer_run.py에서 설명)
            토, 일은 수집하지 않으나 공휴일, 휴일의 경우 동작한다. 이부분은 다른 연동 작업을 통한 개선 필요
            
    Line 11-23: OneDrive(저장을 위한 장소, 다른 cloud 또는 폴더로 대체 가능)와 chromedriver(chrome crolling을 위한) 경로 설정 필요
    
    Line 25
    
 2. 
