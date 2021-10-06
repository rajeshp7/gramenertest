from datetime import datetime
from time import sleep
import yaml
import os
import time
import json
from os.path import dirname
from json2html import json2html
import logging
import sys
import pyodbc
from threading import Thread
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as expCond
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementNotVisibleException
from selenium.common.exceptions import ElementNotSelectableException
from selenium.common.exceptions import ElementClickInterceptedException


# log configuration
# logging.basicConfig(filename='debug.log',
#                     level=logging.DEBUG, format='%(asctime)s %(message)s')
logging.basicConfig(stream=sys.stderr, level=logging.INFO,
                    format='%(asctime)s %(message)s')
# read config file


def read_yaml(config_file_path):
    yaml_contents = yaml.load(
        open(config_file_path, encoding="utf8"), Loader=yaml.FullLoader)
    return yaml_contents


class Database:
    def __init__(self):
        pass

    def connect_to_database(self, db_credentials):
        db_name = db_credentials['database']
        if(db_name == 'mysql'):
            # db_connection = self.mysql_db_connection(db_credentials)
            # TO DO
            pass
        elif(db_name == 'mssql'):
            db_connection = self.mssql_db_connection(db_credentials)
        else:
            print("database is not correct")
        return db_connection

    def mssql_db_connection(self, db_credentials):
        db_connector = pyodbc.connect("DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={0};database={1};UID={2};PWD={3}".format(
            db_credentials['server'], db_credentials['db_name'], db_credentials['db_user'], db_credentials['db_password']))
        return db_connector

    def select_query(self, db_credentials, query):
        """Execute a select query, store the value in a dictionary
        and return the result

        Args:
            db_credentials (dic): database server and credentials
            query (str): select query to execute

        Returns:
            str: execution result
            str: execution result description
        """
        result = ""
        global auto_dictionary
        temp_query = query.split('|')
        res_var = temp_query[0]
        query = temp_query[1]
        try:
            db_connection = self.connect_to_database(db_credentials)
            db_cursor = db_connection.cursor()
            db_cursor.execute(query)
            row = [list(i) for i in db_cursor.fetchall()]
            if len(row) == 1:
                t_row = row[0]
                auto_dictionary[res_var] = t_row[0]
            elif len(row) > 1:
                temp_row_val = []
                for j in range(0, len(row)):
                    temp_row_val.append("".join(row[j]))
                auto_dictionary[res_var] = temp_row_val
            result = "pass"
            result_description = "Select query successful"
        except Exception as e:
            row = 'null'
            result = "fail"
            result_description = "Select query failed due to "+str(e)
        finally:
            db_connection.close()
        return result, result_description


class Actions:
    """Class which holds all the actions required to perform in a test
    """

    def __init__(self):
        pass

    def launch_browser(self, browser):
        """Function to launch a specific browser

        Args:
            browser (str): Browser type string

        Returns:
            object: driver
        """
        global browser_mode
        dir_path = dirname(dirname(os.getcwd()))
        chromedriver_path = dir_path+r'\\drivers\\chromedriver.exe'
        geckodriver_path = dir_path+r'\drivers\geckodriver.exe'

        try:
            if browser.lower() == 'chrome':
                options = webdriver.ChromeOptions()
                options.add_argument("start-maximized")
                options.add_argument("disable-infobars")
                options.add_argument("--disable-extensions")
                options.add_argument("--disable-popup-blocking")
                options.add_argument("--profile-directory=Default")
                options.add_argument("--ignore-certificate-errors")
                options.add_argument("--disable-plugins-discovery")
                options.add_argument("--incognito")
                options.add_argument("--disable-web-security")
                options.add_argument("--allow-running-insecure-content")
                options.add_argument("user_agent=DN")
                options.add_argument("--disable-gpu")
                options.add_experimental_option(
                    'excludeSwitches', ['enable-logging'])
                if browser_mode == 'headless':
                    logging.info("[INFO] : Launching Browser in headless mode")
                    options.add_argument("--headless")
                driver = webdriver.Chrome(options=options,
                                          executable_path=chromedriver_path)
            elif browser.lower() == 'firefox':
                options = webdriver.FirefoxOptions()
                options.add_argument("start-maximized")
                if browser_mode == 'headless':
                    options.add_argument("--headless")
                driver = webdriver.Firefox(options=options,
                                           executable_path=geckodriver_path)
                # TODO
            elif browser.lower() == 'ie':
                pass
                # TODO
            else:
                pass
                # TODO
            driver.delete_all_cookies()
            logging.info("Browser launched : "+browser)
            return driver
        except Exception as e:
            logging.error(e)
            print("Exception launch browser: ", e)

    def switch_window(self, driver):
        """Function to switch to the tab in the browser

        Args:
            driver (Object): webdriver object
            test_data(str): Tab number to switch to
        """
        try:
            current_window = driver.current_window_handle
            child_windows = driver.window_handles
            for window in child_windows:
                if(window != current_window):
                    driver.switch_to.window(window)
                    break
            result = "pass"
            result_description = "Switch to window action successful"
        except Exception as e:
            print('Exception occured while input: ', e)
            result = "fail"
            result_description = "Switch to window action failed due to an exception " + \
                str(e)
        return result, result_description

    def input_value(self, _element, test_data):
        """Function to input a value in a text field

        Args:
            driver (object): webdriver object
            obj_prop (str): object property

        Returns:
            str: result of the input action
        """
        try:
            _element.clear()
            _element.send_keys(test_data)
            result = "pass"
            result_description = "Input action successful"
        except Exception as e:
            print('Exception occured while input: ', e)
            result = "fail"
            result_description = "Input action failed due to an exception " + \
                str(e)
        return result, result_description

    def click(self, driver, _element):
        """Function to click on an element

        Args:
            _element (object): Web element

        Returns:
            str: result of click action
        """
        try:
            _element.click()
            result = "pass"
            result_description = "Click action successful"
        except Exception as e:
            logging.error(e)
            result = "fail"
            result_description = "Click action failed due to an exception " + \
                str(e)
        return result, result_description

    def wait(self, test_data):
        """Function to wait for a period of time

        Args:
            test_data (str): time to wait

        Returns:
            str: Result of wait action
        """
        try:
            sleep(int(test_data))
            result = "pass"
            result_description = "wait action successful"
        except Exception as e:
            logging.error(e)
            print('Exception occured while wait: ', e)
            result = "fail"
            result_description = "Wait action failed due to an exception " + \
                str(e)
        return result, result_description

    def text(self, _element, test_data):
        """Function to verify the text of an element

        Args:
            _element (obj): web element
            test_data (str): value to compare with

        Returns:
            str: result of the text function
        """
        try:
            temp_text = _element.text
            if(temp_text == test_data):
                result = "pass"
                result_description = "text verification successful"
            else:
                print("comparison failed, expected %s and actual %s" %
                      (test_data, temp_text))
                result = "fail"
                result_description = "text verification failed," + \
                    test_data+" expected but "+temp_text+" is actual"
        except Exception as e:
            logging.error(e)
            print("Exception in text", e)
            result = "fail"
            result_description = "text verification failed " +\
                "due to exception" + str(e)
        return result, result_description

    def attribute(self, _element, test_data):
        """ Added get attribute method to get the attribute value from that locator """

        """Function to get the attribute of an element which helps to verify the presence of 
        visuals/ no data images etc.

        Args:
            _element (obj): web element
            test_data (str): locator|attribute value

        Returns:
            str: result of the attribute function
        """
        try:
            temp_var = test_data.split('|')
            attr_method = temp_var[0]
            temp_value = temp_var[1]
            attr = _element.get_attribute(attr_method)
            if(attr == temp_value):
                result = "pass"
                result_description = "verification successful"
            else:
                print("comparision failed, expected %s and actual %s" %
                      (temp_value, attr))
                result = "fail"
                result_description = "verification failed," + \
                    temp_value+" expected but "+attr+" is actual"
        except Exception as e:
            logging.error(e)
            print("Exception in verification", e)
            result = "fail"
            result_description = "verification failed " +\
                "due to exception" + str(e)
        return result, result_description

    def close(self, driver):
        """Function to close the browser

        Args:
            driver (obj): web driver

        Returns:
            str: Result of the close browser action
        """
        try:
            driver.quit()
            result = "pass"
            result_description = "Close browser successful"
        except Exception as e:
            logging.error(e)
            print("Exception while closing the browser: ", e)
            result = "fail"
            result_description = "Close failed due to an exception " + \
                str(e)
        return result, result_description

    def expression(self, expression):
        """Methos to execute an expression

        Args:
            expression (str): expression details

        Returns:
            str: action result and description
        """
        try:
            global auto_dictionary
            temp_exp = expression.split('|')
            store_var = temp_exp[0]
            exp_type = temp_exp[1]
            result_description = "Expression excuted successfully"
            if exp_type.lower() != 'per':
                temp_params = temp_exp[2].split(',')
                param1 = int(auto_dictionary[temp_params[0]])
                param2 = int(auto_dictionary[temp_params[1]])
            if exp_type.lower().startswith('add'):
                auto_dictionary[store_var] = param1 + param2
                result = "pass"
            elif exp_type.lower().startswith('sub'):
                auto_dictionary[store_var] = param1 - param2
                result = "pass"
            elif exp_type.lower().startswith('mul'):
                auto_dictionary[store_var] = param1 * param2
                result = "pass"
            elif exp_type.lower().startswith('div'):
                auto_dictionary[store_var] = param1 / param2
                result = "pass"
            elif exp_type.lower().startswith('per'):
                auto_dictionary[store_var] = auto_dictionary[temp_exp[2]] * 100
                result = "pass"
            elif exp_type.lower().startswith('eql'):
                if param1 == param2:
                    result = "pass"
                else:
                    result = "fail"
            else:
                print('expression not correct')
                result = "fail"
                result_description = "Provided expression is not valid"
        except Exception as e:
            print('Exception occured while executing expression', e)
            result = "fail"
            result_description = "expression failed due to an exception " + \
                str(e)

        return result, result_description

    def select_option(self, _element, test_data):
        """Select an option from dropdown

        Args:
            driver (Object): Webdriver object
            _element (web element):
            element on which select action must be performed
            test_data (str): option details

        Returns:
            str: result of the action
        """
        try:
            select = Select(_element)
            temp_var = test_data.split('|')
            select_type = temp_var[0]
            select_value = temp_var[1]
            if select_type == 'index':
                select.select_by_index(select_value)
            elif select_type == 'text':
                select.select_by_visible_text(select_value)
            elif select_type == 'value':
                select.select_by_value(select_value)
            result = "pass"
            result_description = "Select action successful"
        except Exception as e:
            print("Exception occurred in select ", e)
            result = "fail"
            result_description = "Select failed due to an exception " + \
                str(e)
        return result, result_description

    def dropdown_options(self, _element, test_data):
        """Validate the options available in the dropdown

        Args:
            _element (Web element): drop down element
            test_data (list): list of the options to be validated

        Returns:
            str : result of the action
        """
        result = "pass"
        try:
            # print(test_data)
            select = Select(_element)
            options = select.options
            options_text = []
            for i in range(0, len(options)):
                options_text.append(options[i].text)
            for j in range(0, len(test_data)):
                result_description = "dropdown options validation successful"
                if not test_data[j] in options_text:
                    print(" option does not exist ", test_data[j])
                    result = "fail"
                    result_description = "option does not exist " + \
                        test_data[j]
        except Exception as e:
            print("Exception occurred in dropdown_options ", e)
            result = "fail"
            result_description = "validation failed due to an exception" + \
                str(e)
        return result, result_description

    def placeholder(self, _element, test_data):
        """Validate the placeholder text

        Args:
            _element (webelement): element for which
            placeholder should be validated
            test_data (str): expected value

        Returns:
            str: action result and result description
        """
        try:
            placeholder_text = _element.get_attribute('placeholder')
            if placeholder_text == test_data:
                result = "pass"
                result_description = "placeholder text validation successful"
            else:
                result = "fail"
                result_description = "placeholder text not matching " +\
                    test_data + "is expected but " + placeholder_text +\
                    " is actual"
        except Exception as e:
            result = "fail"
            result_description = "Exception in placeholder validation " +\
                str(e)
        return result, result_description

    def hover(self, driver, _element):
        """hover pointer to an element

        Args:
            driver (object): webdriver
            _element (webelement): Web element

        Returns:
            str : action result and description
        """
        try:
            ActionChains(driver).move_to_element(_element).perform()
            result = "pass"
            result_description = "hover to the element is successful"
        except Exception as e:
            result = "fail"
            result_description = "hover failed due to an exception " + \
                str(e)
        return result, result_description

    def switch_to_frame(self, driver, _element):
        """Swith to frame

        Args:
            driver (object): webdriver
            _element (webelement): Web element

        Returns:
            str: action result and description
        """
        try:
            driver.switch_to.frame(_element)
            result = "pass"
            result_description = "Switch to frame successfully"
        except Exception as e:
            result = "fail"
            result_description = "Switch to frame failed due to "+str(e)
        return result, result_description

    def color(self, _element, test_data):
        """Verify the color code of the element

        Args:
            _element (webelement): web element
            test_data (str): color code

        Returns:
            str: action result and description
        """
        try:
            act_code = _element.get_attribute('fill')
            if act_code == test_data:
                result = "pass"
                result_description = "color match successful"
            else:
                result = "fail"
                result_description =\
                    "color match failed expected "+test_data +\
                    " actual "+act_code
        except Exception as e:
            result = "failed"
            result_description =\
                "color match failed due to an exception "+str(e)
        return result, result_description

    def alert_text(self, driver, test_data):
        """Validate the alert text

        Args:
            driver (obj): webdriver
            test_data (str): Expected text

        Returns:
            str: action result and description
        """
        try:
            alert_text = driver.switch_to_alert().text()
            if alert_text == test_data:
                result = "pass"
                result_description = "alert text validation successful"
            else:
                result = "fail"
                result_description = "alert text expected " + test_data +\
                    " actual " + alert_text
        except Exception as e:
            result = "fail"
            result_description = "alert_text failed due to "+str(e)
        return result, result_description

    def alert_close(self, driver):
        """[summary]

        Args:
            driver ([type]): [description]

        Returns:
            [type]: [description]
        """
        try:
            driver.switch_to_alert().accept()
            result = "pass"
            result_description = "alert close successful"
        except Exception as e:
            result = "fail"
            result_description = "alert close failed due to "+str(e)
        return result, result_description

    def scroll_page(self, driver):
        try:
            driver.execute_script(
                "window.scrollBy(0,document.body.scrollHeight)")
            result = "pass"
            result_description = "Scroll action successful"
        except Exception as e:
            result = "fail"
            result_description = "scroll action failed due to "+str(e)
        return result, result_description

    def script(self, driver, test_data):
        """Execute JavaScript

        Args:
            driver (obj): webdriver
            test_data (str): script to be executed

        Returns:
            str : action result and description
        """
        global auto_dictionary
        try:
            if "|" in test_data:
                temp_test_data = test_data.split("|")
                var = temp_test_data[0]
                script = temp_test_data[1]
                auto_dictionary[var] = driver.execute_script(script)
            else:
                driver.execute_script(test_data)
            result = "pass"
            result_description = "Script execution successful"
        except Exception as e:
            result = "fail"
            result_description = "Script execution failed due to "+str(e)
        return result, result_description

    def comparelist(self, test_data):
        """Method to compare a list of elements

        Args:
            test_data (str): lists to compare

        Returns:
            str: action result and description
        """
        global auto_dictionary
        try:
            temp_test_var = test_data.split(",")
            temp_list1 = auto_dictionary[temp_test_var[0]]
            temp_list2 = auto_dictionary[temp_test_var[1]]
            if temp_list1 == temp_list2:
                result = "pass"
                result_description = "values are equal"
            else:
                result = "fail"
                result_description = "values are not equal"
        except Exception as e:
            result = "fail"
            result_description = "failed due to an exception "+str(e)
        return result, result_description


class Start_Execution(Actions):
    """Class to start test script execution

    Args:
        Actions (class): Actions class
    """

    def __init__(self, _paths, browser, db_credentials):
        self._paths = _paths
        self.browser = browser
        self.db_credentials = db_credentials

    # Connect to database
    database = Database()

    # Collect Test Scripts

    def read_yaml(self, file_path):
        """Function to read yaml file

        Args:
            file_path (str): path to the yaml file

        Returns:
            obj: contents of the yaml file
        """
        content = yaml.load(open(file_path, encoding="utf8"),
                            Loader=yaml.FullLoader)
        return content

    def collect_test_scripts(self, test_scripts_fpath):
        """Function to collect all the test scripts in a folder

        Args:
            test_scripts_fpath (str): test scripts folder path

        Returns:
            list: list of test scripts
        """
        _scripts = os.listdir(test_scripts_fpath)
        return _scripts

    # Read Test Script
    def read_test_script(self, test_script_path):
        """Function to read the test script

        Args:
            test_script_path (str): path of the test script

        Returns:
            str: test script name
            str: test steps
        """
        actions = yaml.load(
            open(test_script_path, encoding="utf8"), Loader=yaml.FullLoader)
        name = actions['name']
        steps = actions['steps']
        return name, steps

    # Method to convert timedelta to string
    def timeConverter(self, totalTime):
        if isinstance(totalTime, datetime.timedelta):
            return totalTime.total_seconds()

    # Wrtie test_script result
    def write_test_script_result(self, filename, test_scripts_result):
        """Wite test results to Yaml file

        Args:
            filename (str): path to create file and file name
            test_scripts_result (obj): test scripts results object
        """
        res_json = json.loads(json.dumps(
            test_scripts_result, default=self.timeConverter))
        format_to_table = json2html.convert(json=res_json)
        with open(filename+'.yaml', 'w') as result:
            yaml.dump(test_scripts_result, result)
        with open(filename+'.html', 'w') as htmlresult:
            htmlresult.write(format_to_table)
            htmlresult.close()

    # Get Test Object property

    def get_object_property(self, obj_name):
        """Function to get the object property from the object repo

        Args:
            obj_name (str): object name with combination of screen name
            and control name. eg: Login-username

        Returns:
            str: object property
        """
        temp_obj = obj_name.split('-')
        page_name = temp_obj[0]
        control_name = temp_obj[1]
        obj_repo_filePath = self._paths['test_objects_path']
        test_objects = self.read_yaml(obj_repo_filePath)
        obj_prop = ''
        for test_object in test_objects:
            if test_object.lower() == page_name.lower():
                obj_controls = test_objects[test_object]
                for obj_control in obj_controls:
                    try:
                        if (obj_control[control_name]):
                            obj_prop = obj_control[control_name]
                            return obj_prop
                    except Exception:
                        pass
        return obj_prop

    # Get Test Data
    def get_test_data(self, dataVar, _data):
        """Function to get get_test_data

        Args:
            dataVar (list): list of test data
            _data (str) : variable of test data

        Returns:
            str: variable value
        """
        global auto_dictionary
        try:
            if _data in dataVar:
                test_data = dataVar.get(_data)
                if str(test_data).startswith('~'):
                    test_data = auto_dictionary[test_data.strip('~')]
            else:
                print(_data + " not available in test data file")
                test_data = 'NA'
        except Exception as e:
            test_data = 'NA'
            print(_data + "not available", e)
        return test_data

    # Get Test Element

    def get_test_element(self, driver, test_element):
        """Function to get the test element property

        Args:
            driver (obj): webdriver
            test_element (str): web element on which an action to be performed

        Returns:
            str: element
        """
        global loader
        try:
            test_element = test_element.strip('\'')
            # driver.implicitly_wait(3)
            driver_wait = WebDriverWait(driver, 100, poll_frequency=10,
                                        ignored_exceptions=[
                                            ElementNotVisibleException,
                                            ElementNotSelectableException,
                                            ElementClickInterceptedException])
            if loader != 'NA':
                driver_wait.until(
                    expCond.invisibility_of_element_located(
                        (By.CLASS_NAME, loader)))
            if test_element.startswith('#'):
                test_element = test_element.lstrip('#')
                element = driver_wait.until(
                    expCond.presence_of_element_located((By.ID, test_element)))
            elif test_element.startswith(('//', '/')):
                element = driver_wait.until(
                    expCond.presence_of_element_located(
                        (By.XPATH, test_element)))
            elif test_element.startswith('.'):
                test_element = test_element.lstrip('.')
                element = driver_wait.until(
                    expCond.presence_of_element_located
                    ((By.CLASS_NAME, test_element)))
            elif test_element.startswith('tag'):
                test_element = test_element.split('.')
                element = driver_wait.until(
                    expCond.presence_of_element_located
                    ((By.TAG_NAME, test_element[1])))
            elif test_element.startswith('>'):
                test_element = test_element.lstrip('>')
                element = driver_wait.until(
                    expCond.presence_of_element_located((By.CSS_SELECTOR, test_element)))
            elif test_element.startswith('>'):
                test_element = test_element.lstrip('>')
                element = driver_wait.until(
                    expCond.presence_of_element_located((By.CSS_SELECTOR, test_element)))
            else:
                element = driver_wait.until(
                    expCond.presence_of_element_located
                    ((By.LINK_TEXT, test_element)))
        except Exception as e:
            element = ""
            print("Exception occurred in get_test_element: ", test_element, e)
        return element

    def open_application(self, driver, app_url):
        """Function to open the application in the browser

        Args:
            driver (obj): webdriver
            app_url (str): application url

        Returns:
            str: result of the function execution
        """
        try:
            driver.maximize_window()
            driver.get(app_url)
            result = "pass"
            result_description = "Application launched successfully"
        except Exception as e:
            print("Exception occurred while opening application: ", e)
            result = "fail"
            result_description = "Launch application failed due to " + str(e)
        return result, result_description

    # Execute test step
    def execute_test_step(self, driver, action, test_element, test_data):
        """ Function to execute a test step

        Args:
            driver (obj): Webdriver
            action (str): Action to perform on a test element
            test_element (str): Element on which an action to be performed
            test_data (str): test data required while performing an action

        Returns:
            [type]: [description]
        """
        try:
            action_result = ""
            res_desc = ""
            action = action.lower()
            if action == "launch":
                action_result, res_desc = self.open_application(
                    driver, test_data)
            elif action == "switchwindow":
                action_result, res_desc = self.switch_window(
                    driver)
            elif action == "input":
                action_result, res_desc = self.input_value(
                    test_element, test_data)
            elif action == "click":
                action_result, res_desc = self.click(driver, test_element)
            elif action == "wait":
                action_result, res_desc = self.wait(test_data)
            elif action == "text":
                action_result, res_desc = self.text(test_element, test_data)
            elif action == "attribute":
                action_result, res_desc = self.attribute(
                    test_element, test_data)
            elif action == "close":
                action_result, res_desc = self.close(driver)
            elif action == "selectquery":
                action_result, res_desc = self.database.select_query(
                    self.db_credentials, test_data)
            elif action == "expression":
                action_result, res_desc = self.expression(test_data)
            elif action == "select":
                action_result, res_desc = self.select_option(
                    test_element, test_data)
            elif action == "dropdownoptions":
                action_result, res_desc = self.dropdown_options(
                    test_element, test_data)
            elif action == "placeholder":
                action_result, res_desc = self.placeholder(
                    test_element, test_data)
            elif action == "hover":
                action_result, res_desc = self.hover(driver,
                                                     test_element)
            elif action == "color":
                action_result, res_desc = self.color(test_element, test_data)
            elif action == "scrollpage":
                action_result, res_desc = self.scroll_page(driver)
            elif action == "alerttext":
                action_result, res_desc = self.alert_text(driver, test_data)
            elif action == "alertclose":
                action_result, res_desc = self.alert_close(driver)
            elif action == "switchtoframe":
                action_result, res_desc = self.switch_to_frame(test_element)
            elif action == "script":
                action_result, res_desc = self.script(driver, test_data)
            elif action == "comparelist":
                action_result, res_desc = self.comparelist(test_data)
            else:
                print("Action does not exist, please check ", action)
                action_result = "fail"
                res_desc = "Action does not exist"
        except Exception as e:
            print(action+": error occured in execute_test_step: ", e)
            action_result = "fail"
            res_desc = "error occured in execute_test_step: "+str(e)
        finally:
            return action_result, res_desc

    # Script Execution
    def scripts_execution(self):
        """Function to execute all the scripts
        available in test scripts folder.
            It is based on the execution mode.
        """
        try:
            logging.info(
                "[INFO] : Executing all the scripts \
                available in the scripts folder")
            # collect Test Scripts
            test_scripts = self.collect_test_scripts(
                self._paths['test_scripts_path'])
            logging.info("[INFO] : Total " +
                         str(len(test_scripts))+" scripts to execute")
            test_scripts_result = []
            # test results file name
            res_filename = r'\results_' + \
                str(datetime.now())+'_'+self.browser
            res_filename = res_filename.replace(":", "-")
            test_results_path = self._paths['test_results_path'] + res_filename
            for test_script in test_scripts:
                startTime = time.time()
                # Test Script results
                test_script_result = self.start_execution(test_script)
                endTime = time.time()
                test_script_result['Duration(sec)'] = round(
                    (endTime - startTime), 2)
                test_scripts_result.append(test_script_result)
        except Exception as e:
            logging.error(e)
            print("Exception occured in Scripts execution ", e)
        finally:
            # write test results to yaml file
            self.write_test_script_result(
                test_results_path, test_scripts_result)
            logging.info(
                "[INFO] : Scripts Execution Completed, Check The Results")
    # Start Execution

    def start_execution(self, _script):
        """Function to strat test script execution

        Args:
            _script_path (str): test script file path

        Returns:
            obj: test script execution result
        """
        try:
            test_script_filepath = self._paths['test_scripts_path'] + \
                '/' + _script
            test_data_filepath = self._paths['test_data_path'] + \
                '/' + _script
            test_script_name, test_steps = self. read_test_script(
                test_script_filepath)
            test_data_dump = self.read_yaml(test_data_filepath)
            test_script_result = {
                "Test Script": test_script_name,
                "Result": "",
                "Duration(sec)": "",
                "Test_Steps_Result": []
            }
            logging.info("[INFO] : Executing Script "+_script)
            for testdata in test_data_dump:
                driver = self.launch_browser(self.browser)
                # Loop through steps
                test_steps_result = []
                res_flag = 0
                for test_step in test_steps:
                    logging.info("[INFO] : Executing Step: "+test_step)
                    step_startTime = time.time()
                    split_step = test_step.split(" ")
                    action = split_step[0]
                    if split_step[1] != 'NA':
                        try:
                            test_element = self.get_object_property(
                                split_step[1])
                            _element = self.get_test_element(
                                driver, test_element)
                        except Exception as e:
                            logging.error(e)
                            logging.info("[ERROR] : "+str(e))
                            print("Exception in object "+split_step[1])
                    else:
                        _element = 'NA'
                    if split_step[2] != 'NA':
                        try:
                            test_data = self.get_test_data(
                                testdata, split_step[2])
                        except Exception as e:
                            logging.error(e)
                            logging.info("[ERROR] : "+str(e))
                            print("Exception in test data "+split_step[2])
                    else:
                        test_data = 'NA'
                    test_step_result, test_step_result_desc = \
                        self.execute_test_step(
                            driver, action, _element, test_data)
                    step_endTime = time.time()
                    if(test_step_result.lower() == "fail"):
                        res_flag = 1
                    temp_result = {
                        "Test Step": test_step,
                        "Test Step Result": test_step_result,
                        "Test Result Description": test_step_result_desc,
                        "duration(sec)":
                        round((step_endTime - step_startTime), 2)}
                    test_steps_result.append(temp_result)
                if res_flag == 1:
                    test_script_result["Result"] = "Fail"
                else:
                    test_script_result["Result"] = 'Pass'
        except Exception as e:
            logging.error(e)
            test_script_result["Result"] = "Fail"
            print("Exception occurred in start execution ", e)
        finally:
            driver.quit()
            test_script_result['Test_Steps_Result'].append(
                test_steps_result)
            # test_script_result['Test_Steps_Result'] = test_steps_result
            return test_script_result

    def suite_execution(self):
        """Function to execute a test suite, based on the execution mode
        """
        try:
            test_suites = self.read_yaml(self._paths['test_suite_path'])
            test_suites_result = []
            suite_res_flag = 0
            # test results file name
            res_filename = r'\suite_results_' + \
                str(datetime.now())+'_'+self.browser
            res_filename = res_filename.replace(":", "-")
            test_results_path = self._paths['test_results_path'] + res_filename
            # test suite execution
            for suite in test_suites:
                test_suite_result = {
                    "Suite": suite['Suite Name'], "Result": '',
                    "Duration(sec)": '', "Test Scripts Result": []}
                suite_start_time = time.time()
                execute_status = suite['Execute']
                if execute_status.lower() == 'yes':
                    test_scripts = suite['Test Scripts']
                    test_scripts_result = []
                    for script in test_scripts:
                        script_name = script['Name']
                        script_execution_status = script['Execute']
                        if script_execution_status.lower() == 'yes':
                            startTime = time.time()
                            test_script = script_name+'.yaml'
                            test_script_result = self.start_execution(
                                test_script)
                            endTime = time.time()
                            test_script_result['Duration(sec)'] = \
                                round(endTime - startTime, 2)
                            test_scripts_result.append(test_script_result)
                            if(test_script_result['Result'].lower() == "Fail"):
                                suite_res_flag = 1
                        else:
                            script['Result'] = "Skip"
                            script['Duration(sec)'] = 0
                            test_scripts_result.append(script)
                    suite_end_time = time.time()
                    if(suite_res_flag == 1):
                        test_suite_result['Result'] = 'Pass'
                    else:
                        test_suite_result['Result'] = 'Fail'
                    test_suite_result['Duration(sec)'] = \
                        round(suite_end_time - suite_start_time, 2)
                    test_suite_result['Test Scripts Result'] = \
                        test_scripts_result
                    test_suites_result.append(test_suite_result)

                else:
                    test_suite_result['Result'] = "Skip"
                    test_suite_result['Duration(sec)'] = 0
                    test_suites_result.append(test_suite_result)
        except Exception as e:
            logging.error(e)
            print("Exception occured in test suite execution ", e)
        finally:
            self.write_test_script_result(
                test_results_path, test_suites_result)
            logging.info("[INFO] : Suite Execution Completed, Check Results")


if __name__ == "__main__":

    auto_dictionary = {}

    dir_path = dirname(dirname(os.getcwd()))

    home_path = dirname(dirname(dirname(os.getcwd())))

    # folder paths
    test_scripts_path = home_path+r'\artefacts\scripts'
    test_results_path = home_path+r'\artefacts\results'
    test_objects_path = home_path+r'\artefacts\objects'
    test_data_path = home_path+r'\artefacts\data'
    test_suite_path = home_path+r'\artefacts\suite'

    # config_file_path
    config_file_path = home_path+r'\artefacts\config.yaml'

    exec_params = read_yaml(config_file_path)

    # Browsers
    browser = exec_params['browser_details']['browser']
    browser_mode = exec_params['browser_details']['mode']

    # loader class
    loader = exec_params['loader_class']

    # Execution mode
    execution_mode = exec_params['execution_mode']

    # create test scripts results directory
    try:
        os.makedirs(test_results_path)
    except OSError:
        print("results directory exists")

    folder_paths = {'test_scripts_path': test_scripts_path,
                    'test_results_path': test_results_path,
                    'test_objects_path': test_objects_path +
                    r'\object_repo.yaml',
                    'test_suite_path': test_suite_path+r'\test_suite.yaml',
                    'test_data_path': test_data_path}

    # db Connection:
    db_details = exec_params['db_details']
    db_credentials = {'database': db_details['database'],
                      'server': db_details['db_server'],
                      'db_name': db_details['db_name'],
                      'db_user': db_details['db_user'],
                      'db_password': db_details['db_password']}

    # Initiate Execution
    def initiate_execution(folder_paths, browser, db_credentials):
        execution = Start_Execution(
            folder_paths, browser, db_credentials)
        try:
            if execution_mode == 1:
                execution.scripts_execution()
            elif execution_mode == 2:
                execution.suite_execution()
            else:
                print("Input Valid Execution mode")
        except Exception as e:
            logging.error(e)
            print("Exception occurred in main execution ", e)

    # If the browser provided in the config file is a list,
    # then execution will start in multiple browsers.
    if type(browser) is list:
        thread_list = []
        for i, browse in enumerate(browser):
            thr = Thread(target=initiate_execution, args=[
                folder_paths, browse, db_credentials])
            thread_list.append(thr)
            thr.start()
        for thr in thread_list:
            thr.join()
    else:
        initiate_execution(folder_paths, browser, db_credentials)
