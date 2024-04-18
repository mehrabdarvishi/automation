from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
import pandas as pd
import time


course_types = {
	'تئوری': 'نظري',
	'عملی': 'عملي',
	'تئوری - عملی': 'عملي نظري'
}


new_course_button_id = 'ctl00_CH1_btnNew'

course_id_input_id = 'ctl00_CH1_frmLessonDefinition_txtLCode'
course_name_input_id = 'ctl00_CH1_frmLessonDefinition_txtLName'
course_theory_credit_input_id = 'ctl00_CH1_frmLessonDefinition_txtNCredit'
course_practical_credit_input_id = 'ctl00_CH1_frmLessonDefinition_txtACredit'

course_type_select_id = 'ctl00_CH1_frmLessonDefinition_ddlLanType'
submit_button_id = 'ctl00_CH1_frmLessonDefinition_btnEdit'


driver = webdriver.Chrome()

wait = WebDriverWait(driver, 10)

url = '😔'

driver.get(url)

driver.add_cookie ({ 'name' : '.ASPXFORMSAUTH', 'value' : '😡'})
driver.add_cookie ({ 'name' : 'ASP.NET_SessionId', 'value' : '😛'})

df = pd.read_excel('courses.xlsx')

for index, row in df.iterrows():

	driver.get(url)
	semester_search_input = driver.find_element(By.ID, new_course_button_id).click()


	wait.until(EC.visibility_of_element_located((By.ID, 'cboxLoadedContent')))

	frame = driver.find_element(By.CLASS_NAME, 'cboxIframe')
	driver.switch_to.frame(frame)



	course_type_select = Select(driver.find_element(By.ID, course_type_select_id))
	course_type = course_types[row['نوع درس']]
	course_type_select.select_by_visible_text(course_type)
	time.sleep(1)

	course_id_input = driver.find_element(By.ID, course_id_input_id)
	course_id_input.send_keys(Keys.CONTROL +  'a')
	course_id_input.send_keys(row['کد درس'])
	time.sleep(1)
	driver.find_element(By.ID, course_name_input_id).send_keys(row['نام درس'])
	if course_type == 'نظري' or course_type == 'عملي نظري':
		time.sleep(1)
		driver.find_element(By.ID, course_theory_credit_input_id).send_keys('1')
	if course_type == 'عملي' or course_type == 'عملي نظري':
		time.sleep(1)
		driver.find_element(By.ID, course_practical_credit_input_id).send_keys('1')
	time.sleep(1)
	driver.find_element(By.ID, submit_button_id).click()

	time.sleep(4)

driver.close()