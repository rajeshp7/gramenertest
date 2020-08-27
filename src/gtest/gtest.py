import yaml
import os
from os.path import dirname
import logging
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as expCond
from selenium.webdriver.common.by import By
from selenium.common.exceptions import ElementNotVisibleException, ElementNotSelectableException, ElementClickInterceptedException
import chromedriver_binary
from time import sleep


# log configuration
logging.basicConfig(filename='execution.log',
                    level=logging.DEBUG, format='%(asctime)s %(message)s')
# format='%(asctime)s %(message)s', level=logging.INFO)
# read config file


def read_yaml(config_file_path):
    yaml_contents = yaml.load(
        open(config_file_path), Loader=yaml.FullLoader)
    return yaml_contents


class Actions:
    def __init__(self):
        pass

    def launch_browser(self, browser):
        """Function to launch a specific browser

        Args:
            browser ([str]): [Browser type string]

        Returns:
            [object]: [driver]
        """
        try:
            if browser.lower() == 'chrome':
                driver = webdriver.Chrome()
                return driver
            elif browser.lower() == 'firefox':
                pass
                # TODO
            elif browser.lower() == 'ie':
                pass
                # TODO
            else:
                pass
                # TODO
            logging.info("Browser launched : ")
        except Exception as e:
            logging.error(e)

    def input_value(self, _element, test_data):
        """Function to input a value in a text field

        Args:
            driver (object): webdriver object
            obj_prop (str): object property

        Returns:
            str: result of the action
        """
        try:
            _element.send_keys(test_data)
            return "pass"
        except Exception as e:
            print('Exception occured while input: ', e)
            return "fail"

    def click(self, _element):
        try:
            # sleep(10)
            _element.click()
            return "pass"
        except Exception as e:
            logging.error(e)
            print('Exception occured while click: ', e)
            return "fail"

    def wait(self, test_data):
        try:
            sleep(int(test_data))
            return "pass"
        except Exception as e:
            logging.error(e)
            print('Exception occured while wait: ', e)
            return "fail"

    def text(self, _element, test_data):
        try:
            temp_text = _element.text
            if(temp_text == test_data):
                return "pass"
            else:
                return "fail"
        except Exception as e:
            logging.error(e)
            print("Exception in text")
            return "fail"


class Start_Execution(Actions):

    def __init__(self, _paths, browser):
        self._paths = _paths
        self.browser = browser

    # Collect Test Scripts

    def collect_test_scripts(self, test_scripts_fpath):
        _scripts = os.listdir(test_scripts_fpath)
        return _scripts

    # Read Test Script
    def read_test_script(self, test_script_path):
        actions = yaml.load(open(test_script_path), Loader=yaml.FullLoader)
        name = actions['name']
        steps = actions['steps']
        return name, steps

    # get Test Element
    def get_test_element(self, driver, test_element):
        try:
            test_element = test_element.strip('\"')
            driver_wait = WebDriverWait(driver, 60, poll_frequency=1, ignored_exceptions=[
                                        ElementNotVisibleException, ElementNotSelectableException, ElementClickInterceptedException])
            if test_element.startswith('#'):
                test_element = test_element.lstrip('#')
                element = driver_wait.until(
                    expCond.presence_of_element_located((By.ID, test_element)))
            elif test_element.startswith(('//', '/')):
                element = driver_wait.until(
                    expCond.element_to_be_clickable(
                        (By.XPATH, test_element)))
            elif test_element.startswith('.'):
                test_element = test_element.lstrip('.')
                element = driver_wait.until(
                    expCond.presence_of_element_located
                    ((By.CLASS_NAME, test_element)))
            else:
                element = driver_wait.until(
                    expCond.presence_of_element_located
                    ((By.CSS_SELECTOR, test_element)))
            return element
        except Exception as e:
            print("Exception occurred in get_test_element: ", test_element, e)

    def open_application(self, driver, app_url):
        try:
            driver.maximize_window()
            driver.get(app_url)
            return "pass"
        except Exception as e:
            print("Exception occurred while opening application: ", e)
            return "fail"

    # Execute test step
    def execute_test_step(self, driver, action, test_element, test_data):
        try:
            action_result = ""
            action = action.lower()
            if action == "launch":
                action_result = self.open_application(driver, test_data)
            elif action == "input":
                action_result = self.input_value(test_element, test_data)
            elif action == "click":
                action_result = self.click(test_element)
            elif action == "wait":
                action_result = self.wait(test_data)
            elif action == "text":
                action_result = self.text(test_element, test_data)
            return action_result
        except Exception as e:
            print("error occured in execute_test_step: ", e)
            return "fail"

    # Start Execution

    def start_execution(self):
        """[summary]
        """
        # collect Test Scripts
        test_scripts = self.collect_test_scripts(
            self._paths['test_scripts_path'])

        # Read each Test Script
        for test_script in test_scripts:
            test_script_path = self._paths['test_scripts_path']+'/'+test_script
            test_script_name, test_steps = self. read_test_script(
                test_script_path)
            driver = self.launch_browser(self.browser)
            # Test Script results
            test_script_result = {"Test Script": test_script_name}
            # Loop through steps
            test_steps_result = []
            for test_step in test_steps:
                res_flag = 0
                split_step = test_step.split(" ")
                action = split_step[0]
                test_element = split_step[1]
                if test_element != 'NA':
                    _element = self.get_test_element(driver, test_element)
                else:
                    _element = 'NA'
                test_data = split_step[2]
                test_step_result = self.execute_test_step(
                    driver, action, _element, test_data)
                if(test_step_result.lower() == "fail"):
                    res_flag = 1
                temp_result = {
                    "Test Step": test_step,
                    "Test Step Result": test_step_result, "duration": "0s"}
                test_steps_result.append(temp_result)

            if res_flag == 1:
                test_script_result['Result'] = "Fail"
            else:
                test_script_result['Result'] = 'Pass'
            print(test_steps_result)
            print(test_script_result)


if __name__ == "__main__":

    # config_file_path
    config_file_path = dirname(dirname(os.getcwd()))+r'\\config.yaml'
    exec_params = read_yaml(config_file_path)

    home_path = dirname(dirname(dirname(os.getcwd())))
    # Browsers
    browser = exec_params['browser']

    # folder paths
    test_scripts_path = home_path+r'\test scripts'
    test_results_path = home_path+r'\test results'

    folder_paths = {'test_scripts_path': test_scripts_path,
                    'test_results_path': test_results_path}

    # Initiate Execution
    execution = Start_Execution(folder_paths, browser)
    execution.start_execution()
