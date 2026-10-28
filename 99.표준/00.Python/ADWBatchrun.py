import psycopg2
import os
import datetime
import configparser
import sys
import logging
import decrypto
import importlib

#logging.basicConfig(level=logging.INFO)

logger = logging.getLogger(' ')
logger.setLevel(level=logging.INFO)

# 공통 변수 셋팅
totlog = ""
step_num = 0
juje = sys.argv[1][4:7] #주제영역셋팅
comment_line = ""
pkresult = 0
table_name = sys.argv[1][4:-3].lower()
# 프로그램 시작
start_time = datetime.datetime.now()  # 시작 시간 저장
logger.warning("=======================================================")
logger.warning("■ 프로그램 시작     : {}".format(start_time.strftime("%Y%m%d%H%M%S")))
logger.warning("■ 프로그램 PGM_ID   : " + sys.argv[1])
totlog += "=======================================================\n■ 프로그램 시작     : {}".format(start_time.strftime("%Y%m%d%H%M%S")) + "\n■ 프로그램 PGM_ID   : " + sys.argv[1] + '\n'

# Parameter 셋팅
if(len(sys.argv) < 3):
    logger.warning("■ ERROR.")
    logger.warning("■ 입력된 Parameter가 정확하지 않습니다.")
    logger.warning("■ USAGE) python BatchRun.py Jobname DATE")
    logger.warning("■    EX) python BatchRun.py UCD_WCDIFS01TM_TG 20230426")
    logger.warning("=======================================================")
    exit()

# 입력 날짜 셋팅
if(len(sys.argv[2]) == 8):
    base_dt = sys.argv[2]
elif(len(sys.argv[2]) == 6):
    base_dt = sys.argv[2] + '01'
else:
    logger.warning("■ ERROR.")
    logger.warning("■ 입력된 날짜형식이 정확하지 않습니다.")
    logger.warning("=======================================================")
    exit()

# 경로 셋팅
body_file_path = r'/home/etladm/BATCH/ORG/' + juje + '/' + sys.argv[1] + '.sql'  # EC2 서버 실행시 1번째 파라미터
gen_file_path = r'/home/etladm/BATCH/GEN/' + juje + '/' + sys.argv[1] + '_' + sys.argv[2] + '.sql'
log_file_name = r'/home/etladm/BATCH/LOG/' + juje + '/log_' + sys.argv[1] + '_' + sys.argv[2] + '.dat'

try:
    # env.ini 파일에서 정보 가져오기
    config = configparser.ConfigParser()
    config.read(r'/home/etladm/BATCH/COM/ENV/anenv.ini')
    host = config['REDSHIFT']['host']
    port = config['REDSHIFT']['port']
    dbname = config['REDSHIFT']['dbname']
    user = config['REDSHIFT']['user']
    password = config['REDSHIFT']['password']

    # db password 복호화
    depassword = decrypto.decrypt_aes(password)

    h_dong_cd = '4413132000'
    # body 파일 읽으며 변수치환
    with open(body_file_path, 'r') as f:
        body_sql = f.read().format(h_dong_cd=h_dong_cd, BASE_DT=base_dt)

    # gen 파일 생성
    with open(gen_file_path, 'w') as f:
        f.write(body_sql)

    # Redshift에 연결
    conn = psycopg2.connect(
        host=host,
        port=port,
        dbname=dbname,
        user=user,
        password=depassword
    )

    sql_file = open(gen_file_path, 'r')
    sql_array = sql_file.read().split('####SQL')  # body 파일내 쿼리들을 순차적으로 배열에 저장, .format 으로 변수설정
    sql_file.close()

    # total_count = 0 # 누적 처리 건수 초기화
    with conn.cursor() as cur:
        for sql in sql_array:
            # comment_line = sql.split('\n')[0] # sql 첫번째 문장은 comment_line
            for line in sql.split('\n'):
                if "P_BAT_STG_NM" in line:
                    start_index = sql.find("P_BAT_STG_NM :") + 15
                    end_index = sql.find("  */", start_index)
                    comment_line = sql[start_index:end_index].strip()
            if (sql.strip() != ""):
                try:
                    cur.execute(sql)
                    count = cur.rowcount  # 처리 건수
                    # total_count += count # 누적 처리 건수 업데이트
                    logger.warning(
                        ("■ " + comment_line.replace('-- ', '')).ljust(19, ' ') + f" : {count} rows affected")
                    totlog += "■ " + comment_line.replace('-- ', '').ljust(17, ' ') + f" : {count} rows affected\n"
                    conn.commit()
                except Exception as e:
                    if "can't execute an empty query" in str(e):  # 해당 에러발생시 작업 재수행(주석걸린 쿼리 수행시 발생하는 에러)
                        continue
                    else:
                        # logger.warning(f"Error occurred: {e}")
                        # logger.warning(f"SQL statement: {sql}")
                        conn.rollback()
                        end_time = datetime.datetime.now()
                        duration = end_time - start_time
                        logger.warning("=======================================================")
                        logger.warning("                     <프로그램 오류>")
                        logger.warning("■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")))
                        logger.warning("■ 프로그램 실행시간 : {}".format(duration))
                        logger.warning("■ LOG_FILE_PATH     : " + log_file_name)
                        logger.warning("=======================================================")
                        totlog += "=======================================================\n"
                        totlog += "                     <프로그램 오류>\n"
                        totlog += "■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")) + "\n"
                        totlog += "■ 프로그램 실행시간 : {}".format(duration) + "\n"
                        totlog += f"Error occurred: {e}\n"
                        totlog += f"SQL statement: {sql}\n"
                        totlog += "=======================================================\n"
                        with open(log_file_name, 'w') as f:
                            f.write(totlog)
                        exit(1)

    # 변경 내용 저장
    conn.commit()

    pkcur = conn.cursor()

    # 프로시저 실행
    pkcur.execute('CALL public.sp_table_key_chk(%s)', (table_name,))

    # 결과 가져오기 (필요한 경우)
    pkresult = pkcur.fetchall()
    pkresult = pkresult[0][0]
    if (pkresult == 0):
        logger.warning("■ PK_Consistency    : TRUE")
        totlog += "■ PK_Consistency    : TRUE\n"
    else:
        conn.rollback()
        end_time = datetime.datetime.now()
        duration = end_time - start_time
        logger.warning("■ PK_Consistency    : FALSE")
        logger.warning("=======================================================")
        logger.warning("                     <프로그램 오류>")
        logger.warning("■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")))
        logger.warning("■ 프로그램 실행시간 : {}".format(duration))
        logger.warning("■ PK 중복 Error")
        logger.warning("=======================================================")
        totlog += "■ PK_Consistency    : FALSE\n"
        totlog += "=======================================================\n"
        totlog += "                     <프로그램 오류>\n"
        totlog += "■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")) + "\n"
        totlog += "■ 프로그램 실행시간 : {}".format(duration) + "\n"
        totlog += "■ PK 중복 Error\n"
        totlog += "=======================================================\n"
        with open(log_file_name, 'w') as f:
            f.write(totlog)
        exit(1)
    # 커밋
    conn.commit()
except Exception as e:
    conn.rollback()
    end_time = datetime.datetime.now()
    duration = end_time - start_time
    logger.warning("=======================================================")
    logger.warning("                     <프로그램 오류>")
    logger.warning("■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")))
    logger.warning("■ 프로그램 실행시간 : {}".format(duration))
    logger.warning(f"Error occurred: {e}\n")
    logger.warning("■ LOG_FILE_PATH     : " + log_file_name)
    logger.warning("=======================================================")
    totlog += "=======================================================\n"
    totlog += "                     <프로그램 오류>\n"
    totlog += "■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")) + "\n"
    totlog += "■ 프로그램 실행시간 : {}".format(duration) + "\n"
    totlog += f"Error occurred: {e}\n"
    totlog += "=======================================================\n"
    with open(log_file_name, 'w') as f:
        f.write(totlog)
    exit(1)
finally:
    # 커넥션 닫기
    cur.close()
    pkcur.close()
    conn.close()

    end_time = datetime.datetime.now()  # 종료 시간 저장
    logger.warning("■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")))
    totlog += "■ 프로그램 종료     : {}".format(end_time.strftime("%Y%m%d%H%M%S")) + "\n"

    # 실행 시간 계산
    duration = end_time - start_time
    logger.warning("■ 프로그램 실행시간 : {}".format(duration))
    logger.warning("■ LOG_FILE_PATH     : " + log_file_name)
    logger.warning("=======================================================")
    totlog += "■ 프로그램 실행시간 : {}".format(duration) + "\n=======================================================\n"

    # 로그 생성
    with open(log_file_name, 'w') as f:
        f.write(totlog)
