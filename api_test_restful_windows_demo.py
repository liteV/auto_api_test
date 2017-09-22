#coding=utf-8
# -----------------------------
# Author: WangWenbang         |
# Co.: daokoudai.Beijing      |
# Version: 3.8.0              |
# Date: 6/16/2017             |
# -----------------------------
# instruction:
# 1. you can put the program script in each platform: Windows or Linux or Unix
# 2. configure the path and tag before using it
# 3. configure your email SMTP server , chose the right username and password and the sending destination
# 4. Must use the excel as your case file, fill each line full with no blank! or it will stop during your action
# 5. Better to clean your testing result file in path before run again

import urllib
import urllib2
import json
import hashlib
import md5
import os
import sys
import time
import xlrd
import xlwt
from xlutils.copy import copy
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
import smtplib
import thread
import threading
from threading import Thread
import Queue
import logging
import zipfile

# global arguments
THREAD_NUM = 0
COUNT_FINISH_TAG = 0
COUNT_TOTAL_DICT_LIST = []
FILE_PATH = 'xxxxx'
CASE_END_TAG = '.xlsx'
TEST_RESULT_TAG = '.md'
TEST_RESULT_ZIP_TAG = '.zip'
STARTTIME = time.ctime()

URL_HEAD_CONFIG = 'xxxxxx'


NAME_SMTP = 'xxxxx'
NAME_SENDER_EMAIL = 'xxxxx'
PASSWORD_SENDER_EMAIL = 'xxxxx'
NAME_DEST_EMAIL = 'xxxxx'
TIME_STAMP = '%Y%m%d%H%M'

def tool_get_case_from_file(case_name_):
    data = xlrd.open_workbook(case_name_)
# select sheet from the excel
    table = data.sheet_by_name(u'Sheet1')
    return table

def get_http_reponse(url_list):
    response_res = []
    global URL_HEAD_CONFIG
    for i in url_list:
        url_send = URL_HEAD_CONFIG + str(i)
        try:
            request = urllib2.Request(url_send)
            response_data = urllib2.urlopen(request)
            response_res.append(response_data.read())
        except urllib2.HTTPError,e:
            print "Error code:",e.code
            if e.code == '404':
                print "Page not found!"
                logging.warning('Page not found!')
                continue
            elif e.code == '403':
                print "Access denied!"
                logging.error('Access denied!')
                continue
            elif e.code == '505':
                print "blank in url with no turned into %20 !"
                logging.error('blank in url with no turned into %20 !')
                continue
            else:
                print "Something happened! Error code",e.code
                logging.error('something error happened') 
                continue	
    return response_res

def json_alter(response_list):
# return rescode and resmsg
    return_resCode_list = []
    #return_resMsg_list = []
    #res_tuple_list = []

# res list for resCode
    for i in response_list:
        json_res_resCode = json.loads(i)
        try:
            if "success" in json_res_resCode:
                return_resCode_list.append(("%(success)s"%json_res_resCode).encode('utf-8'))
            elif "resCode" in json_res_resCode:
                return_resCode_list.append(("%(resCode)s"%json_res_resCode).encode('utf-8'))
        except:
            print "tag cannot found...."
            logging.error('tag cannot found')

    return return_resCode_list


def Case_list_get(filepath_, *args_end):
# adapt to more args.
    tags = args_end
    file_list = os.listdir(filepath_)
    file_list_select = []
    for i in file_list:
        if i.endswith(tags):
            file_list_select.append(i)
        else:
            continue
    return file_list_select

def Case_run(t,case_name):
# initial param
    list_id = []
    list_url = []
    list_expect = []

# initial param final store
    list_id_str = []
    list_url_str = []
    list_expect_str = []

# get params from excel
    list_id = tool_get_case_from_file(case_name).col_values(0)
    list_url = tool_get_case_from_file(case_name).col_values(1)
    list_expect = tool_get_case_from_file(case_name).col_values(2)

    for i in range(1, len(list_id)):
        list_id_str.append(str(int(list_id[i])))

    for i in range(1, len(list_url)):
        list_url_str.append(str(list_url[i]))

    for i in range(1, len(list_expect)):
        list_expect_str.append(str(list_expect[i]))

    charge_response_res = []
    charge_response_res = get_http_reponse(list_url_str)

    tag_find_res_tuple_list = []
    tag_find_res_tuple_list = json_alter(charge_response_res)

    return tag_find_res_tuple_list, list_id_str, list_expect_str, charge_response_res, list_url_str

def Case_result_get(tag_tuple_list_in, list_id_tag_, list_expect_tag_, charge_response_res_, list_url_str_request_, case_file_):
# name the testing result file
    global TEST_RESULT_TAG
    global CASE_END_TAG
    global TIME_STAMP
    time_stamp = time.strftime(TIME_STAMP)
    filename = case_file_.replace(CASE_END_TAG, '_')+time_stamp+TEST_RESULT_TAG
# init count content
    count_pass = 0
    count_fail = 0
    count_total_dict = {}
    count_fail_id = []
    global COUNT_TOTAL_DICT_LIST
    global COUNT_FINISH_TAG
    try:
        f_ = open(filename, 'a+')

        res_tuple_to_list = []
        for i in range(len(tag_tuple_list_in)):
            res_tuple_to_list.append(tag_tuple_list_in[i])

        list_id_tag = []
        for i in range(len(list_id_tag_)):
            list_id_tag.append(list_id_tag_[i])	

        list_expect_tag = []
        for i in range(len(list_expect_tag_)):
            list_expect_tag.append(list_expect_tag_[i])

        charge_response_res = []
        for i in range(len(charge_response_res_)):
            charge_response_res.append(charge_response_res_[i])

        list_url_str_request = []
        for i in range(len(list_url_str_request_)):
            list_url_str_request.append(list_url_str_request_[i])

# initial splite list
        #splite_resCode_list =[]
        #splite_resMsg_list = []
        tag_resCode_out = []

# push each expected content to check:
# 1.resCode
# 2.resMsg
        #for i in range(len(res_tuple_to_list)):
            #print res_tuple_to_list[i]
            #splite_resCode_list.append(res_tuple_to_list[i])
            #splite_resMsg_list.append(res_tuple_to_list[i][1])

# make tag for test case
        for i in range(len(res_tuple_to_list)):
            #print isinstance(res_tuple_to_list[i], str)
            if list_expect_tag[i] == '0.0':
                list_expect_tag[i] = '0000'
			
            if res_tuple_to_list[i] == list_expect_tag[i]:
                tag_resCode_out.append("pass")
            elif res_tuple_to_list[i] == 'True': 
                tag_resCode_out.append("pass")
            else:
                tag_resCode_out.append("fail")

# write result into file
        for i in range(len(res_tuple_to_list)):
            if tag_resCode_out[i] == "pass":
                print >> f_, list_id_tag[i],
                print >> f_, tag_resCode_out[i]
                # if pass, still print
                print >> f_, list_url_str_request[i]
                print >> f_, charge_response_res[i]

                count_pass = count_pass+1
            else:
                print >> f_, list_id_tag[i]
                print >> f_, tag_resCode_out[i]
                print >> f_, list_url_str_request[i]
                print >> f_, charge_response_res[i]				
                count_fail_id.append(list_id_tag[i])
                count_fail = count_fail+1

        print >> f_, '\n'
        print >> f_, 'test—result:'
        count_total_dict['pass_num'] = count_pass
        count_total_dict['fail_num'] = count_fail
        print >> f_, count_total_dict
        print >> f_, "failed case id:"
        print >> f_, count_fail_id

        if count_fail>0:
            summary_result = str(filename)+"::"+str(count_total_dict)+"--"+"failed case id:"+str(count_fail_id)
            print >> f_, summary_result
            logging.critical('find fail case test!') 
        elif count_fail==0:
            summary_result = str(filename)+"::"+str(count_total_dict)+"--"+"no failed case id:"
            print >> f_, summary_result
            logging.info('no fail case test OK!') 
    except:
        print 'file write error or list append has happend'
        logging.error('file write error or list append has happend') 
    finally:
        f_.close()
# add summary count in global store space
    COUNT_TOTAL_DICT_LIST.append(count_total_dict)
    COUNT_FINISH_TAG += 1
    print COUNT_FINISH_TAG
    #while COUNT_FINISH_TAG != THREAD_NUM:
	

def Single_Run(t, case_file):
# run the testing case,send http request and acquire the response and the result expected
# 1.res_tag_find_res_tuple_list
# 2.res_list_id_str
# 3.res_list_expect_str
# 4.res_http_response
# 5.res_list_url_str
    res_tag_find_res_tuple_list, res_list_id_str, res_list_expect_str, res_http_response, res_list_url_str = Case_run(t,case_file)
    logging.info('INFO:case run OK!')
# form all the testing result file in the same path the case-file store
    Case_result_get(res_tag_find_res_tuple_list, res_list_id_str, res_list_expect_str, res_http_response, res_list_url_str, case_file)
    logging.info('INFO:test result get OK!')

def Multi_Run(case_list):
    case_num_thread = len(case_list)
    global THREAD_NUM
    THREAD_NUM = case_num_thread
# thread pool
    queue_list = []
    global COUNT_FINISH_TAG
    global STARTTIME
# method and argument ready to add into thread
# 1.method name add to 'target'
# 2.arguments add to 'args'
    start_time = time.ctime()
    print '----------------------'
    STARTTIME = start_time
    print 'case running...\n'
    q = Queue.Queue(case_num_thread)
    for i in range(case_num_thread):
        t = threading.Thread(target=Single_Run, args=(q, case_list[i]))
        queue_list.append(t)
        t.start()
        q.put(t)
    for i in range(case_num_thread):
        queue_list[i].join()


def Result_list_get(filepath_, *args_end):
    tags = args_end
# get all file from the file path we chose
# if the file tag match with the selected, then append to array
    file_list = os.listdir(filepath_)
    file_list_select = []
    for i in file_list:
        if i.endswith(tags):
            file_list_select.append(i)
        else:
            continue
    logging.info('result file ready OK!')

    return file_list_select

def Result_list_zip_get(filepath_, *args_end):
    tags = args_end
    global TIME_STAMP
    global TEST_RESULT_TAG
    global TEST_RESULT_ZIP_TAG

    time_stamp = time.strftime(TIME_STAMP)
    zip_name = 'Test_result_file_'+time_stamp+'.zip'
	
    try:
        f_zip = zipfile.ZipFile(zip_name,'w',zipfile.ZIP_DEFLATED)
        file_list = os.listdir(filepath_)
        for i in file_list:
            if i.endswith(TEST_RESULT_TAG):
                f_zip.write(i)
            else:
                continue
        logging.info('result file write in zip OK!')
    except:
        logging.error('zip file write error!')
    finally:
        f_zip.close()

    list_zip = []
    file_list_new = os.listdir(filepath_)
    for j in file_list_new:
        if j.endswith(TEST_RESULT_ZIP_TAG):
            list_zip.append(j)
    logging.info('zip list OK!')

    return list_zip

def Result_count(res_count_list_):
    num_pass = 0
    num_fail = 0

# summary the num of pass or fail in each testing result file
    for i in res_count_list_:
        num_pass += i['pass_num']
        num_fail += i['fail_num']
    print 'pass num:', num_pass
    print 'fail num:', num_fail
    logging.info('INFO:count calculate result OK!') 
    return num_pass, num_fail

def send_email_init():
    global NAME_SMTP
    global NAME_SENDER_EMAIL
    global PASSWORD_SENDER_EMAIL
    global NAME_DEST_EMAIL
    name_smtp_ = NAME_SMTP
    name_source_mail_ = NAME_SENDER_EMAIL
    name_source_mail_password_ = PASSWORD_SENDER_EMAIL
    name_dest_mail_ = NAME_DEST_EMAIL
    return name_smtp_, name_source_mail_, name_source_mail_password_, name_dest_mail_

def send_email_body_msg_summary(test_result_file_list):
    global FILE_PATH
    email_body = []

# get tail-2 lines from each testing-result file:tag start with 0 as key
    for i in test_result_file_list:
        complete_file_path = FILE_PATH + "\\" + str(i)
        with open(complete_file_path) as f:
            txt = f.readlines()
        keys = [r for r in range(1, len(txt) + 1)]
        result = {k: v for k, v in zip(keys, txt[::-1])}
        total_msg = result[1]
        #print(total_msg)

# merge summary message to make the array
        email_body.append(total_msg)

    return email_body

def send_email_config(sender_server_, sender_mail_name_, sender_password_, reciever_mail_name_, result_list_, title_):
# initial param
    send_server = {'name':'','user':'','password':''}
    send_server['name'] = sender_server_
    send_server['user'] = sender_mail_name_
    send_server['password'] = sender_password_

# sender name to input
    sender_mail_name = sender_mail_name_

# destination ready to send : can be list
    reciever_mail_name = [reciever_mail_name_]

    #files = [logdir+'/tmp_process.log', logdir+'/tmp_keyword_match.log', logdir+'tmp_print.log', logdir+'tmp_autocdp.log', logdir+'net_speed.log', logdir+'alive_stat.log']
    result_files_list = result_list_
# mail title
    mail_title = title_
    return send_server, sender_mail_name, reciever_mail_name, result_files_list, mail_title

def Send_email_action(server, from_, to_, subject_title_, text, files=[]):
    assert type(server) == dict 
    assert type(to_) == list 
    assert type(files) == list

    msg = MIMEMultipart()
    msg['From'] = from_ 
    msg['Subject'] = subject_title_
    msg['To'] = COMMASPACE.join(to_)        # COMMASPACE==', '
    msg['Date'] = formatdate(localtime=True) 
    msg.attach(MIMEText(text)) 

    for file in files:
        part = MIMEBase('application', 'octet-stream')  # 'octet-stream': binary data
        part.set_payload(open(file, 'rb').read()) 
        encoders.encode_base64(part) 
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(file)) 
        msg.attach(part)

    # TLS encrypted in use，for SMTP object made
    smtp = smtplib.SMTP(server['name'])
    smtp.starttls()
    smtp.login(server['user'], server['password']) 
    smtp.sendmail(from_,to_,msg.as_string()) 
    smtp.close()

def Send_email(name_smtp_, name_source_mail_, name_source_mail_password_, name_dest_mail_, result_list, name_subject_, name_body_msg_):
# initial send main config
# we can get the below ones after run this method
# 1.send_server
# 2.send mail name
# 3.send receiver's mail name
# 4.mail title which include the num of pass and fail
    send_server, sender_mail_name, receiver_mail_name, files_res, mail_title = send_email_config(name_smtp_,name_source_mail_,name_source_mail_password_,name_dest_mail_,result_list,name_subject_)
    try:

# make email body massage
# each testing-result file get tail-2 lines for summary
        body = ""
        for i in name_body_msg_:
            body = body + str(i) + "\n"

# send email using arguments with:
# 1.send_server
# 2.send mail name
# 3.send receiver's mail name
# 4.mail title which include the num of pass and fail
# 5.mail body which include summary of each testing-result file
# 6.testing-result file include as attachments
        Send_email_action(send_server, sender_mail_name, receiver_mail_name, mail_title, body, files_res)
        logging.info('INFO:send message by email OK!')
    except:
# if sending email failed,then give out the message
        print 'smtp server or send error happend'
        logging.error('smtp server or send error happend') 

if __name__ == "__main__":
    global FILE_PATH
    global CASE_END_TAG
    global TEST_RESULT_TAG
    global COUNT_FINISH_TAG

    time_stamp = time.strftime('%Y-%m-%d-%H-%M-%S')
    logging.basicConfig(level=logging.DEBUG,  
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',  
                        datefmt='%a,%d%b%Y%H:%M:%S',  
                        filename='./api_test_'+time_stamp+'.log',  
                        filemode='w')

# get chosen tag file from your selected store path, run and get the case list
    case_list = Case_list_get(FILE_PATH, CASE_END_TAG)

# start multi-thread run each case
    Multi_Run(case_list)
    
    print 'start at:', STARTTIME
    end_time = time.ctime()
    print 'end at:', end_time
    print '----case run over----'

    print 'COUNT_FINISH_TAG:', COUNT_FINISH_TAG

    if COUNT_FINISH_TAG == THREAD_NUM:
        num_pass, num_fail = Result_count(COUNT_TOTAL_DICT_LIST)
        print 'waiting...'
# all selected file which include the testing results
        result_list = Result_list_get(FILE_PATH, TEST_RESULT_TAG)
        result_list_zip = Result_list_zip_get(FILE_PATH, TEST_RESULT_ZIP_TAG)

        email_body_msg = send_email_body_msg_summary(result_list)
        #print email_body_msg

# init the configure message
        name_smtp_, name_source_mail_, name_source_mail_password_, name_dest_mail_ = send_email_init()
        name_subject_ = 'Testing result-Pass:'+str(num_pass)+' Fail:'+str(num_fail)

# send email with the upper arguments
        Send_email(name_smtp_, name_source_mail_, name_source_mail_password_, name_dest_mail_, result_list_zip, name_subject_, email_body_msg)
        print '----email send over----'
    else:
        print 'merge failed'
        logging.error('merge failed') 

    























































