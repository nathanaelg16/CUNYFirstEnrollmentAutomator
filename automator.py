from datetime import datetime
from datetime import timedelta
from mailthon import postman, email
from selenium import webdriver
import selenium.common.exceptions
import mysql.connector as mysql
import configparser
import pgpy

config = configparser.ConfigParser()
config.read('config.properties')


def main():
    db = config['DATABASE']

    cnx = mysql.connect(user=db['user'], password=db['password'], host=db['host'], database=db['database'])
    cursor = cnx.cursor()

    find_first = "SELECT username, enrollment_appt FROM EnrollmentQueue ORDER BY enrollment_appt LIMIT 1;"
    query = "SELECT college, term, encrypted_password, email, tries " \
            "FROM EnrollmentQueue WHERE username = %s AND enrollment_appt = %s;"

    cursor.execute(find_first)
    for (username, enrollment_appt) in cursor:
        if datetime.now() >= enrollment_appt:
            cursor.execute(query, (username, enrollment_appt))
            for (college, abbv_term, enc_password, email_addr, tries) in cursor:
                term = get_term_name(abbv_term)
                password = decipher_password(enc_password, config['PGP']['secret'])
                reg_success = perform_registration(username, password, college, term)
                user_info = {
                    'email': email_addr,
                    'username': username,
                    'college': college,
                    'term': abbv_term,
                    'enrollment_appt': enrollment_appt,
                    'tries': tries
                }
                on_complete(reg_success, user_info, cursor)
    cnx.commit()
    cnx.close()


def get_term_name(abbv_term):
    term = abbv_term[2:6]
    if abbv_term[0:2] == "FA":
        term = term + " Fall Term"
    elif abbv_term[0:2] == "SP":
        term = term + " Spring Term"
    elif abbv_term[0:2] == "SU":
        term = term + " Summer Term"
    return term


def decipher_password(blob, secret):
    key, _ = pgpy.PGPKey.from_file('CFAutomator.asc')
    with key.unlock(secret) as ukey:
        enc_pass = pgpy.PGPMessage.from_blob(blob)
        password = ukey.decrypt(enc_pass)
    return password.message.decode(encoding="utf-8").rstrip()


def perform_registration(username, password, college, term):
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)

    try:
        driver.get("http://cf.bbcuny.ml")
        if "CUNY Login" not in driver.title:
            elem = driver.find_element_by_link_text("SIGN OUT")
            elem.click()
        elem = driver.find_element_by_name("usernameH")
        elem.clear()
        elem.send_keys(username + "@login.cuny.edu")
        elem = driver.find_element_by_name("password")
        elem.clear()
        elem.send_keys(password)
        driver.find_element_by_name("submit").click()
        if driver.title != "Employee-facing registry content":
            raise Exception
        driver.find_element_by_link_text("Student Center").click()
        try:
            driver.find_element_by_id("DERIVED_SCC_SUM_PERSON_NAME")
        except selenium.common.exceptions.NoSuchElementException:
            try:
                driver.find_element_by_link_text("Student Center").click()
            except selenium.common.exceptions.NoSuchElementException:
                raise Exception
        driver.switch_to.frame("TargetContent")
        driver.execute_script("javascript:submitAction_win0(document.win0,'DERIVED_SSS_SCR_SSS_LINK_ANCHOR3');")
        elems = driver.find_elements_by_xpath(
            "/html/body/form/div[5]/table/tbody/tr/td/div/table/tbody/tr[8]/td[2]/div/table/tbody/tr")
        found = False
        for elm in elems:
            if elm.get_attribute("id").find("trSSR_DUMMY_RECV1$0_row") != -1:
                if elm.find_element_by_xpath(".//td[2]/div/span").text == term:
                    if elm.find_element_by_xpath(".//td[4]/div/span").text == college:
                        found = True
                        elm.find_element_by_xpath(".//td[1]/div/input").click()
                        driver.execute_script(
                            "javascript:submitAction_win0(document.win0,'DERIVED_SSS_SCT_SSR_PB_GO');")
                        break
        if not found:
            raise Exception
        driver.find_element_by_name("DERIVED_REGFRM1_LINK_ADD_ENRL$82$").click()
        elems = driver.find_elements_by_xpath(
            "/html/body/form/div[5]/table/tbody/tr/td/div/table/tbody/tr[11]/td[2]/div/table/tbody/tr")
        for elm in elems:
            if elm.get_attribute("id").find("trSSR_SS_ERD_ER$0_row") != -1:
                if elm.find_element_by_xpath(".//td[3]/div/div/img").get_attribute("alt") == "Error":
                    raise Exception
        driver.close()
        return True
    except Exception:
        driver.close()
        return False


def on_complete(success, user_info, cursor):
    if not success:
        tries = user_info['tries']
        if tries < 1:
            tries = tries + 1
            new_date = user_info['enrollment_appt'] + timedelta(minutes=10)
            query = "UPDATE EnrollmentQueue " \
                    "SET tries = %s, enrollment_appt = %s " \
                    "WHERE username = %s AND college = %s AND term = %s;"
            cursor.execute(query, (
            tries, new_date.strftime('%Y-%m-%d %H:%M:%S'), user_info['username'], user_info['college'],
            user_info['term']))
            return
    query = "DELETE FROM EnrollmentQueue WHERE username = %s AND college = %s AND term = %s;"
    cursor.execute(query, (user_info['username'], user_info['college'], user_info['term']))
    send_email(success, user_info)


def send_email(success, user_info):
    content = 'Your CUNYFirst Automated Enrollment for the ' + user_info['term'] + ' at ' + user_info[
        'college'] + ' has completed '
    if success:
        status = 'Success'
        content = content + 'successfully.'
    else:
        status = 'Failure'
        content = content + 'with one or more errors on 2 attempts. ' \
                            'Please check CUNYFirst to fix your errors and finish your enrollment.'
    email_config = config['EMAIL']
    p = postman(host=email_config['email_host'], port=email_config['port'],
                auth=(email_config['email'], email_config['email_pass']), force_tls=True)
    e = email(sender='Nathanael Gutierrez <nathanaelg16@gmail.com>',
              subject='CUNYFirst Enrollment Automator: ' + status,
              content=content,
              encoding='utf8',
              receivers=[user_info['email']])
    p.send(e)


if __name__ == '__main__':
    main()
