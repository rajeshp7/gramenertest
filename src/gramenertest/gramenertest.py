from datetime import datetime
from time import sleep
import yaml
import os
import time
from os.path import dirname
import logging
import pyodbc
import mysql.connector
from mysql.connector import Error
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as expCond
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementNotVisibleException
from selenium.common.exceptions import ElementNotSelectableException
from selenium.common.exceptions import ElementClickInterceptedException


# log configuration
logging.basicConfig(filename='debug.log',
                    level=logging.DEBUG, format='%(asctime)s %(message)s')

# read config file


def read_yaml(config_file_path):
    yaml_contents = yaml.load(
        open(config_file_path), Loader=yaml.FullLoader)
    return yaml_contents


class Database:
    def __init__(self):
        pass

    def mysql_db_connection(self, db_credentials):
        db_connector = mysql.connector.connect(
            host=db_credentials['host'], database=db_credentials['db_name'], user=db_credentials['user'], password=db_credentials['password'])
        return db_connector

    def mssql_db_connection(self, db_credentials):
        db_connector = pyodbc.connect("DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={0}; database={1}; UID={2};PWD={3}".format(
            db_credentials['server'], db_credentials['db_name'], db_credentials['db_user'], db_credentials['db_password']))
        return db_connector

    def connect_to_database(self, db_credentials):
        db_name = db_credentials['database']
        if(db_name == 'mysql'):
            db_connection = self.mysql_db_connection(db_credentials)
        elif(db_name == 'mssql'):
            db_connection = self.mssql_db_connection(db_credentials)
        else:
            print("database is not correct")
        return db_connection

    def execute_select_query(self, db_credentials, query):
        global auto_dictionary
        temp_query = query.split('|')
        res_var = temp_query[0]
        query = temp_query[1]
        try:
            db_connection = self.connect_to_database(db_credentials)
            db_cursor = db_connection.cursor()
            db_cursor.execute(query)
            row = db_cursor.fetchone()
            temp_row = str(row).split(',')
            if '(' in temp_row[0]:
                row = temp_row[0].replace('(', '')
            auto_dictionary[res_var] = row
            result = "pass"
        except Exception as e:
            row = 'null'
            result = "fail"
            print("Exception occurred while executing a query ", e)
        return result


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
        dir_path = dirname(dirname(os.getcwd()))
        chromedriver_path = dir_path+r'\\drivers\\chromedriver.exe'
        geckodriver_path = dir_path+r'\drivers\geckodriver.exe'
        try:
            if browser.lower() == 'chrome':
                driver = webdriver.Chrome(
                    executable_path=chromedriver_path)
            elif browser.lower() == 'firefox':
                driver = webdriver.Firefox(
                    executable_path=geckodriver_path)
                # TODO
            elif browser.lower() == 'ie':
                pass
                # TODO
            else:
                pass
                # TODO
            driver.delete_all_cookies()
            logging.info("Browser launched : ")
            return driver
        except Exception as e:
            logging.error(e)
            print("Exception launch browser: ", e)

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
            return "pass"
        except Exception as e:
            print('Exception occured while input: ', e)
            return "fail"

    def click(self, driver, _element):
        """Function to click on an element

        Args:
            _element (object): Web element

        Returns:
            str: result of click action
        """
        try:
            _element.click()
            return "pass"
        except Exception as e:
            logging.error(e)
            print('Exception occured while click: ', e)
            return "fail"

    def wait(self, test_data):
        """Function to wait for a period of time

        Args:
            test_data (str): time to wait

        Returns:
            str: Result of wait action
        """
        try:
            sleep(int(test_data))
            return "pass"
        except Exception as e:
            logging.error(e)
            print('Exception occured while wait: ', e)
            return "fail"

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
            print(temp_text)
            if(temp_text == test_data):
                return "pass"
            else:
                print("comparision failed, expected %s and actual %s" %
                      (test_data, temp_text))
                return "fail"
        except Exception as e:
            logging.error(e)
            print("Exception in text", e)
            return "fail"

    def close(self, driver):
        """Function to close the browser

        Args:
            driver (obj): web driver

        Returns:
            str: Result of the close browser action
        """
        try:
            driver.quit()
            return "pass"
        except Exception as e:
            logging.error(e)
            print("Exception whil closeing the browser: ", e)
            return "fail"

    def execute_expression(self, expression):
        try:
            global auto_dictionary
            temp_exp = expression.split('|')
            store_var = temp_exp[0]
            exp_type = temp_exp[1]
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

            else:
                print('expression not correct')
                result = "fail"
        except Exception as e:
            print('Exception occured while executing expression', e)
            result = "fail"

        return result

    def select_option(self, _element, test_data):
        """Select an option from dropdown

        Args:
            driver (Object): Webdriver object
            _element (web element): element on which select action must be performed
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
        except Exception as e:
            print("Exception occurred in select ", e)
            result = "fail"
        return result

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
                if not test_data[j] in options_text:
                    print(" option does not exist ", test_data[j])
                    result = "fail"
        except Exception as e:
            print("Exception occurred in dropdown_options ", e)
            result = "fail"
        return result


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
        content = yaml.load(open(file_path),
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
        actions = yaml.load(open(test_script_path), Loader=yaml.FullLoader)
        name = actions['name']
        steps = actions['steps']
        return name, steps

    # Wrtie test_script result
    def write_test_script_result(self, filename, test_scripts_result):
        """Wite test results to Yaml file

        Args:
            filename (str): path to create file and file name
            test_scripts_result (obj): test scripts results object
        """
        with open(filename, 'w') as result:
            yaml.dump(test_scripts_result, result)

    # Get Test Object property
    def get_object_property(self, obj_name):
        """Function to get the object property from the object repo

        Args:
            obj_name (str): object name with combination of screen name and control name. eg: Login-username

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
        if _data in dataVar:
            test_data = dataVar.get(_data)
            if str(test_data).startswith('~'):
                test_data = auto_dictionary[test_data.strip('~')]
        else:
            print("Test data is not available")
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
        try:
            test_element = test_element.strip('\'')
            driver_wait = WebDriverWait(driver, 100, poll_frequency=1,
                                        ignored_exceptions=[
                                            ElementNotVisibleException,
                                            ElementNotSelectableException,
                                            ElementClickInterceptedException])
            driver_wait.until(
                expCond.invisibility_of_element_located(
                    (By.CLASS_NAME, "loading")))
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
            else:
                element = driver_wait.until(
                    expCond.presence_of_element_located
                    ((By.LINK_TEXT, test_element)))
            return element
        except Exception as e:
            print("Exception occurred in get_test_element: ", test_element, e)

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
            return "pass"
        except Exception as e:
            print("Exception occurred while opening application: ", e)
            return "fail"

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
            action = action.lower()
            if action == "launch":
                action_result = self.open_application(driver, test_data)
            elif action == "input":
                action_result = self.input_value(test_element, test_data)
            elif action == "click":
                action_result = self.click(driver, test_element)
            elif action == "wait":
                action_result = self.wait(test_data)
            elif action == "text":
                action_result = self.text(test_element, test_data)
            elif action == "close":
                action_result = self.close(driver)
            elif action == "executeselectquery":
                action_result = self.database.execute_select_query(self.db_credentials,
                                                                   test_data)
            elif action == "executeexpression":
                action_result = self.execute_expression(test_data)
            elif action == "select":
                action_result = self.select_option(test_element, test_data)
            elif action == "dropdown_options":
                action_result = self.dropdown_options(test_element, test_data)
            else:
                print("Action does not exsit, please check")
                action_result = "fail"
            return action_result
        except Exception as e:
            print("error occured in execute_test_step: ", e)
            return "fail"

    # Script Execution
    def scripts_execution(self):
        """Function to execute all the scripts available in test scripts folder.
            It is based on the execution mode.
        """
        try:
            # collect Test Scripts
            test_scripts = self.collect_test_scripts(
                self._paths['test_scripts_path'])
            test_scripts_result = []
            for test_script in test_scripts:
                startTime = time.time()
                # Test Script results
                test_script_result = self.start_execution(test_script)
                endTime = time.time()
                test_script_result['duration(sec)'] = endTime - startTime
                test_scripts_result.append(test_script_result)
            # test results file name
            res_filename = r'\results_'+str(datetime.now())+'.yaml'
            res_filename = res_filename.replace(":", "-")
            test_results_path = self._paths['test_results_path'] + res_filename
            # write test results to yaml file
            self.write_test_script_result(
                test_results_path, test_scripts_result)
        except Exception as e:
            logging.error(e)
            print("Exception occured in Scripts execution ", e)
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
                "Test Script": test_script_name, "Test_Steps_Result": []}
            for testdata in test_data_dump:
                driver = self.launch_browser(self.browser)
                # Loop through steps
                test_steps_result = []
                res_flag = 0
                for test_step in test_steps:
                    step_startTime = time.time()
                    split_step = test_step.split(" ")
                    action = split_step[0]
                    if split_step[1] != 'NA':
                        test_element = self.get_object_property(
                            split_step[1])
                        _element = self.get_test_element(driver, test_element)
                    else:
                        _element = 'NA'
                    if split_step[2] != 'NA':
                        test_data = self.get_test_data(
                            testdata, split_step[2])
                    else:
                        test_data = 'NA'
                    test_step_result = self.execute_test_step(
                        driver, action, _element, test_data)
                    step_endTime = time.time()
                    if(test_step_result.lower() == "fail"):
                        res_flag = 1
                    temp_result = {
                        "Test Step": test_step,
                        "Test Step Result": test_step_result,
                        "duration(sec)": step_endTime - step_startTime}
                    test_steps_result.append(temp_result)
                test_script_result['Test_Steps_Result'].append(
                    test_steps_result)
                if res_flag == 1:
                    test_script_result['result'] = "Fail"
                else:
                    test_script_result['result'] = 'Pass'
            return test_script_result
        except Exception as e:
            logging.error(e)
            print("Exception occurred in start execution ", e)
        finally:
            driver.quit()

    def suite_execution(self):
        """Function to execute a test suite, based on the execution mode
        """
        try:
            test_suites = self.read_yaml(self._paths['test_suite_path'])
            test_suites_result = []
            suite_res_flag = 0
            for suite in test_suites:
                test_suite_result = {
                    "Suite": suite['Suite Name'], "Result": '', "Duration(sec)": '', "Test Scripts Result": []}
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
                            # test_script = self._paths['test_scripts_path'] + \
                            #     r'\\'+script_name+'.yaml'
                            test_script = script_name+'.yaml'
                            test_script_result = self.start_execution(
                                test_script)
                            endTime = time.time()
                            test_script_result['duration(sec)'] = endTime - \
                                startTime
                            test_scripts_result.append(test_script_result)
                            if(test_script_result['result'].lower() == "Fail"):
                                suite_res_flag = 1

                        else:
                            script['result'] = "Skip"
                            script['duration(sec)'] = 0
                            test_scripts_result.append(script)
                    suite_end_time = time.time()
                    if(suite_res_flag == 1):
                        test_suite_result['result'] = 'Pass'
                    else:
                        test_suite_result['result'] = 'Fail'
                    test_suite_result['duration(sec)'] = suite_end_time - \
                        suite_start_time
                    test_suite_result['Test Scripts Result'].append(
                        test_scripts_result)
                    test_suites_result.append(test_suite_result)

                else:
                    test_suite_result['result'] = "Skip"
                    test_suite_result['duration(sec)'] = 0
                    test_suites_result.append(test_suite_result)
                    # test results file name
            res_filename = r'\suite_results_'+str(datetime.now())+'.yaml'
            res_filename = res_filename.replace(":", "-")
            test_results_path = self._paths['test_results_path'] + res_filename
            # write test results to yaml file
            self.write_test_script_result(
                test_results_path, test_suites_result)
        except Exception as e:
            logging.error(e)
            print("Exception occured in test suite execution ", e)


if __name__ == "__main__":

    auto_dictionary = {}

    dir_path = dirname(dirname(os.getcwd()))

    # config_file_path
    config_file_path = dir_path+r'\\config.yaml'
    exec_params = read_yaml(config_file_path)

    home_path = dirname(dirname(dirname(os.getcwd())))

    # Browsers
    browser = exec_params['browser']

    execution_mode = exec_params['execution_mode']

    # folder paths
    test_scripts_path = home_path+r'\test scripts'
    test_results_path = home_path+r'\test results'
    test_objects_path = home_path+r'\test objects'
    test_data_path = home_path+r'\test data'
    test_suite_path = home_path+r'\test suite'

    # create test scripts results directory
    try:
        os.makedirs(test_results_path)
    except OSError:
        print("results directory exists")

    folder_paths = {'test_scripts_path': test_scripts_path,
                    'test_results_path': test_results_path,
                    'test_objects_path': test_objects_path+r'\test_objects.yaml',
                    'test_suite_path': test_suite_path+r'\test_suite.yaml',
                    'test_data_path': test_data_path}

    # db Connection:
    db_details = exec_params['db_details']
    db_credentials = {'database': db_details['database'], 'server': db_details['db_server'], 'db_name': db_details['db_name'],
                      'db_user': db_details['db_user'], 'db_password': db_details['db_password']}

    # Initiate Execution
    execution = Start_Execution(folder_paths, browser, db_credentials)
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
