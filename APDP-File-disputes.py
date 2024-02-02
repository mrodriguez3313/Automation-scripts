from seleniumbase import Driver
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException


import easygui
import pandas as pd
import numpy as np
import openpyxl
from re import match
import time
import logging
import signal
import sys
import concurrent.futures

# TODO Improvements:
    # 3. Implement back up selectors in case website updates break scripts
    # 4. Propper and standardized logging
    # 7. Create pivot table with the new info and include before and after column.
    # 8. make it into an exe
    # 10. write a testing framework that measures performance (time to file [avg&med]), errors, money saved wrt time, hourly wage dollars/hour,  
    # 12. Switch from Selenium to SeleniumBase
    # 13. Don't overwrite Workbook, so to maintain all other sheets that were already there.
    # 14. Add Comments
    # 15. Refactor code to support higher usability and different scenarios
    # 16. Implement multiprocessing (maybe not)

# TODO bugs:
#   1. check for value in select vendor number first, if its there then fill box, else try to reselect it
#   1a.     Can fix issue by looking for error text instead of forcing X attempts
#   2. Some claims get skipped, get drafted, and don't get refiled if there is a website error that doesn't have the word "Error" in it.
#   6. logging not working properly

def main():
    # Ask User for excel sheet to use
    while True:
        is_error, path = open_file()
        if is_error == None:
            return
        if not is_error:
            break
        

    # Open Chrome in undetectable mode and get login page
    driver = Driver(uc=True)
    driver.get("https://retaillink.login.wal-mart.com/login")

    # Create secondary thread, to load excel sheet and find starting index, while user is logging in.
    with concurrent.futures.ThreadPoolExecutor() as executor:
        future = executor.submit(load_sheet, path)
        sheet_name, df = future.result()

        future2 = executor.submit(catch_up, df)
        index_value = future2.result()
        

    wait_for_user_to_login(driver)
    # instantiate sigInt handler so to not immediately break when window is closed or stopped with Ctrl+C
    signal.signal(signal.SIGINT, signal_handler)
    driver.get("https://retaillink2.wal-mart.com/apdp/CreateDisputes")
    start_time = time.time()

    try:
        loop_over_sheet(driver, df, index_value)
    except Exception as error:
        logging.error("Error while looping over sheet", error)
    finally:
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(elapsed_time / 60, " minutes and ", elapsed_time % 60, " seconds.")
        while True:
            if write_to_file(path, sheet_name, df):
                break
    return 

def loop_over_sheet(driver: WebDriver, df: pd.DataFrame, start_index: int):
    # log of operations
    logging.basicConfig(filename='output.txt', level=logging.DEBUG, format='')
    skip = False
    #  main loop, iterates over every row in the Excel sheet
    for index in range(start_index, len(df)):
        # if a claim's deduction code is 94 AND amount paid is < 0 AND hasn't already been disputed, then file a dispute
        dispute_criteria = (df.at[index, 'DEDUCTION CODE'] == "MERCHANDISE RETURN - DEFECTIVE MERCHANDISE [0094]" and df.at[index, 'Amount Paid($)'] < 0 and (pd.isnull(df.at[index, 'Disputed'])))
        if dispute_criteria:
            claim_start_time = time.time()
            for attempts in range(4):
                if attempts > 2:
                    get_create_dispute_page(driver)
                    click_create_tab(driver)
                    skip = False
                    break

                try:
                    skip, is_err = dispute_process(driver, str(df.at[index, 'Invoice Number']), skip)
                    if is_err:
                        get_create_dispute_page(driver)
                        click_create_tab(driver)
                        continue
                    elif skip == None and is_err == None:
                        df.at[index, 'Disputed'] = "N"
                        continue
                    df.at[index, 'Disputed'] = 'Y'
                    break
                except NoSuchElementException as elemerr:
                    logging.error("Elem error: ", elemerr)
                    skip = False
                    get_create_dispute_page(driver)
                    click_create_tab(driver)
                    continue
                except Exception as error:
                    logging.error("Error:", error)
                    skip = False
                    get_create_dispute_page(driver)
                    click_create_tab(driver)

            claim_end_time = time.time()
            print(f"ttf: {(claim_end_time - claim_start_time):.2f}sec")
    return

def dispute_process(driver: WebDriver, invoice: str, skip: bool) -> (bool, bool):
    is_err = fill_invoice_info(driver, invoice, skip)
    if is_err:
        # error found
        return False, is_err

    disputes_to_file = get_number_of_claims_to_file(driver)

    # Loops over every subclaim found in an invoice number
    for dispute in range(disputes_to_file):
        click_create_dispute(driver, dispute)
        is_err = check_error(driver, "click_create_dispute")
        if is_err:
            # error found
            return False, is_err
        
        if is_disputable(driver):
            is_err = file_dispute(driver)
            if is_err:
                # error found
                return False, is_err 
            click_create_tab(driver)
            skip = False
        elif is_approved(driver):
            click_previous(driver)
            skip = True
            continue
        elif is_draft(driver):
            is_err = file_draft(driver)
            if is_err:
                # error found
                return False, is_err
            click_create_tab(driver)
            skip = False
        elif is_ytbr(driver):
            click_previous(driver)
            skip = True
            continue
        else:
            click_previous(driver)

        # If there are multiple subclaims, after the first is done, re-fill the info to click on next subclaim
        if dispute + 1 < disputes_to_file:
            is_err = fill_invoice_info(driver, invoice, skip)
            if is_err:
                # error found
                return False, is_err

    return skip, is_err

def fill_invoice_info(driver: WebDriver, invoice: str, skip_vendor: bool) -> bool:
    close_notification(driver)
    if not skip_vendor:
        select_vendor_number(driver)

        is_err = check_error(driver, "select_vendor")
        if is_err:
            # error found
            return True

    enter_invoice_number(driver, invoice)

    press_enter(driver)

    is_err = check_error(driver, "press_enter")
    if is_err:
        # error found
        return True
    return False

def is_disputable(driver: WebDriver) -> bool:
    try:
        click_dispute_all(driver)
        return True
    except NoSuchElementException:
        logging.info("DisputeAll button not found")
    except Exception as error:
        logging.error("Error in is_disputable:", error)
    return False

def is_draft(driver: WebDriver) -> bool:
    try:
        driver.find_element(By.XPATH, '//button[contains(@class, "MuiButtonBase-root MuiButton-root MuiButton-contained") and contains(., "DRAFT ALL CANCEL")]')
        return True
    except NoSuchElementException:
        logging.warning("NoSuchElement: Draft all button not found.")
    except Exception as error:
        logging.error("Error in is_draft:", error)
    return False

def is_approved(driver: WebDriver) -> bool:
    try:
        driver.find_element(By.XPATH, '//div[text()="Approved"]')
        return True
    except NoSuchElementException:
        logging.warning("NoSuchElement: 'Approved' text not found")
    except Exception as error:
        logging.error("Error in is_approved:", error)
    return False

def is_ytbr(driver: WebDriver) -> bool:
    try:
        driver.find_element(By.XPATH, '//div[text()="Yet To Be Resolved"]')
        return True
    except NoSuchElementException:
        logging.warning("NoSuchElement: 'ytbr' text not found.")
    except Exception as error:
        logging.error("Error in is_ytbr:", error)
    return False

def file_dispute(driver: WebDriver) -> bool:
    fill_description(driver)

    is_err = check_error(driver, "fill_desc")
    if is_err:
        # error found
        return True
    
    check_description_success(driver)
    
    is_err = submit_sequence(driver)
    if is_err:
        # error found
        return True
    return False

def file_draft(driver: WebDriver) -> bool:
    click_select_lines_box(driver)

    is_err = submit_sequence(driver)
    if is_err:
        # error found
        return True
    return False

def submit_sequence(driver: WebDriver) -> bool:
    click_next(driver)

    is_err = check_error(driver, "click_next")
    if is_err:
        # error found
        return True
    
    click_submit(driver)

    is_err = is_submit_error(driver, "click_submit")
    if is_err:
        # error found
        resubmit(driver)

    click_ok(driver)
    return False

# Specific Actions

def close_notification(driver: WebDriver):
    try:
       WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//a[contains(@class, "MuiButtonBase-root MuiTab-root") and contains(., "Create")]'))).click()
    except TimeoutException:
        logging.warning("TimeoutException: Couldn't find 'create' text in header")
    except Exception as error:
        logging.error("Error in close_notification", error)
    return

def select_vendor_number(driver: WebDriver):
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//button[@title="Open"]'))).click()
        vendor_num = driver.find_element(By.XPATH, '//div[@role="presentation"]')
        vendor_children = vendor_num.find_elements(By.CSS_SELECTOR, "*")
        WebDriverWait(driver, 3).until(EC.element_to_be_clickable(vendor_children[2])).click()
    except TimeoutException as elemerr:
        logging.error("TimeoutException: Dropdown not found", elemerr)
    except Exception as error:
        logging.error("Error in select_vendor_number", error)
    return

def enter_invoice_number(driver: WebDriver, invoice: str):
    print("enter invoice: ", invoice)

    # search for input box with a class MuiInputBase-input MuiInput-input MuiAutocomplete-input MuiAutocomplete-inputFocused MuiInputBase-inputAdornedEnd MuiInputBase-inputMarginDense MuiInput-inputMarginDense
    # if its value is equal to 736533, then continue with entering invoice, else re-enter vendor number either by selecting or by manually inserting text
    invoice_box = ActionChains(driver)
    box = driver.find_element(By.XPATH, '//input[@id="outlined-error-helper-text"]')
    
    while box.get_attribute('value') != '':
        box.send_keys(Keys.BACKSPACE)

    invoice_box.send_keys_to_element(box, invoice)
    invoice_box.perform()
    # ORR after this, check for error text, if error text exists, then reselect or manually enter vendor number.

    i=0
    while box.get_attribute('value') == '':
        driver.sleep(1)
        box = driver.find_element(By.XPATH, '//input[@id="outlined-error-helper-text"]')
        box.click()
        box.send_keys(invoice)
        if i >= 6:
            raise Exception
    return 

def press_enter(driver: WebDriver):
    enter = ActionChains(driver)
    enter.send_keys(Keys.ENTER)
    enter.perform()
    return

def get_number_of_claims_to_file(driver: WebDriver) -> int:
    try:
       return len( driver.find_elements(By.LINK_TEXT, "Create Dispute") )
    except NoSuchElementException:
        logging.debug("Unable to find 'Create Dispute' links")
    except Exception as error:
        logging.error("Error in get_number_of_claims: ", error)

def fill_description(driver: WebDriver):
    description_box = ActionChains(driver)
    box = driver.find_element(By.XPATH, '//textarea[@id="multiline"]')
    description_box.send_keys_to_element(box, "Invalid Reclamation")
    description_box.move_to_element(driver.find_element(By.XPATH, '//button[@class="MuiButtonBase-root MuiButton-root MuiButton-text"]/following-sibling::button'))
    description_box.click()
    description_box.perform()
    return

def check_description_success(driver: WebDriver):
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//p[text()="Successfully Added"]')))
    except TimeoutException as elemerr:
        logging.error("TimeoutException: Succesfully added message not found!", elemerr)
    except Exception as error:
        logging.error("Error in check_description_sucess:", error)
    return

def click_create_dispute(driver: WebDriver, dispute: int):
    disputes = driver.find_elements(By.LINK_TEXT, "Create Dispute")
    disputes[dispute].click()
    return

def click_dispute_all(driver: WebDriver):
    dispute_button = driver.find_element(By.XPATH, '//button[contains(@class, "MuiButtonBase-root MuiButton-root MuiButton-contained") and contains(., "DISPUTE ALL")]')
    dispute_button.click()
    return

def click_select_lines_box(driver: WebDriver):
    select_all_box = ActionChains(driver)
    box = driver.find_element(By.XPATH, '//input[@type="checkbox"]')
    select_all_box.move_to_element(box)
    select_all_box.click()
    select_all_box.send_keys(Keys.ENTER)
    select_all_box.perform()
    return

def click_next(driver: WebDriver):
    try:
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//b[text()="Next"]'))).location_once_scrolled_into_view
    except TimeoutException as elemerr:
        logging.error("TimeoutException: 'Next' button not found", elemerr)
    except Exception as error:
        logging.error("Error in click_next", error)

    next_button = ActionChains(driver)
    button = driver.find_element(By.XPATH, '//b[text()="Next"]')
    next_button.move_to_element(button)
    next_button.click()
    next_button.perform()
    return

def click_previous(driver: WebDriver):
    try:
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//b[text()="Previous"]'))).location_once_scrolled_into_view
    except TimeoutException as elemerr:
        logging.error("TimeoutException: Prev button not found", elemerr)
    except Exception as error:
        logging.error("error in click_previous", error)

    previous_button = ActionChains(driver)
    button = driver.find_element(By.XPATH, '//b[text()="Previous"]')
    previous_button.move_to_element(button)
    previous_button.click()
    previous_button.perform()
    return

def click_prev_alt(driver: WebDriver):
    try:
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//span[text()="Previous"]'))).location_once_scrolled_into_view
    except TimeoutException as elemerr:
        logging.warning("TimeoutException: Prev_b button not found", elemerr)
    except Exception as error:
        logging.error("Error in click_prev_b", error)

    previous_button = ActionChains(driver)
    button = driver.find_element(By.XPATH, '//span[text()="Previous"]')
    previous_button.move_to_element(button)
    previous_button.click()
    previous_button.perform()
    return

def click_submit(driver: WebDriver):
    try:
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//button[contains(@class, "MuiButtonBase-root MuiButton-root MuiButton-contained") and contains(., "Submit")]'))).location_once_scrolled_into_view
    except TimeoutException as elemerr:
        logging.error("TimeoutException: 'Submit' button not found", elemerr)
    except Exception as error:
        logging.error("Error in click_submit", error)

    next_button = ActionChains(driver)
    button = driver.find_element(By.XPATH, '//button[contains(@class, "MuiButtonBase-root MuiButton-root MuiButton-contained") and contains(., "Submit")]')
    next_button.move_to_element(button)
    next_button.click()
    next_button.perform()

    return

def click_ok(driver: WebDriver):
    try:
        WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.XPATH,'//a[contains(@class, "MuiButtonBase-root MuiButton-root MuiButton-text") and contains(.,"Ok")]'))).click()
    except TimeoutException as elemerr:
        logging.warning("TimeoutException: 'Ok' button not found", elemerr)
    except Exception as error:
        logging.error("Error in click_ok", error)
    return

def click_create_tab(driver: WebDriver):
    driver.sleep(.5)
    close_notification(driver)
    return

def close_popup(driver: WebDriver):
    close_button = driver.find_element(By.XPATH, '//button[contains(@aria-label,"close")]')
    close_button.click()
    return

# UTILS

def open_file() -> (bool, str):
    # Returns is_error and path
    try:
        path = easygui.fileopenbox("What Excel sheet would you like to use?")
        if path == None:
            return None, None
        if path.lower().endswith('.xlsx'):
            return False, path
        else:
            easygui.msgbox("Selected file is not an Excel Workbook", "Please pick an Excel file (.xlsx)", "Try Again")
            return True, ""
    except Exception as error:
        logging.error("Error in open_file", error)
        return True, ""
    
def load_sheet(path: str) -> (str, pd.DataFrame):
    while True:
        try:
            xl = pd.ExcelFile(path, engine="openpyxl")
            break
        except PermissionError:
            easygui.msgbox("Selected workbook is in use by another program. Please close workbook before trying again.", "Permission Error", "Try Again")
        
    ws = xl.sheet_names
    sheet_name = list(filter(lambda x: match('^Check_\d{3,10}', x), ws))
    # sheet_name = list(filter(lambda x: match('^Sheet1', x), ws))
    df = xl.parse(sheet_name[0])
    df1 = verify_sheet(df)
    return sheet_name[0], df1

def write_to_file(path: str, sheet_name: str, df: pd.DataFrame) -> bool:
    try:
        with pd.ExcelWriter(path, engine='openpyxl', mode='w') as writer:
            print("PLEASE WAIT! WRITING TO FILE!")
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
    except PermissionError:
        easygui.msgbox("You must close workbook before progress can be saved.", "Permission Error", "Continue")
        return False

def wait_for_response(driver: WebDriver):
    driver.sleep(.5)
    try:
        WebDriverWait(driver, 30).until_not(EC.visibility_of_element_located((By.ID, 'artwork')))
    except Exception as error:
        logging.error("Error in wait_for_response", error)

def check_error(driver: WebDriver, origin: str) -> bool:
# Happens on any request, to fix: Must restart from beginning
    
# Errors that will break code:
# "Service is down, please try after sometime"; happens after click_enter and before click_create_dispute links
    # refresh/restart and then try to re-file
# "Something went wrong!"; can happen at anytime; script will take a really long time to recover or just freeze
    # Needs to hit "ok" button
# "No records found for the search criteria"; Takes a really long time to recover; happens after click_enter and before click_create_dispute
# "Authentication Error"; Forces manual intervention to log out and rerun the script
    wait_for_response(driver)
    try:
        WebDriverWait(driver, .5).until(EC.presence_of_element_located((By.XPATH, '//h6[contains(@class, "MuiTypography-root MuiTypography-h6") and contains(.,"Error")]')))
        print("error found after", origin)
        return True
    except TimeoutException as error:
        logging.info("TimeoutException: No website error found.")
        return False
    except Exception as error:
        logging.error("Error in Check_error", error)
        return False

def is_submit_error(driver: WebDriver, origin: str) -> bool:
    is_err = check_error(driver, origin)
    if is_err:
        close_popup(driver)
        return True
    return False

def resubmit(driver: WebDriver):
    click_prev_alt(driver)

    submit_sequence(driver)
    return

def wait_for_user_to_login(driver: WebDriver):
# Give the user 2 minutes to complete login
    try:
        WebDriverWait(driver, 180).until(EC.title_contains('Retail Link Home'))
    except TimeoutException:
        logging.error("User took too long to enter information or page took too long to load.")
    return

def verify_sheet(df: pd.DataFrame) -> pd.DataFrame:
    # change na values to 0 or '' depending on colum type
    new_df = pd.DataFrame()
    for i, col in enumerate(df):
        if col == "Amount Paid($)" or col == "Invoice Amount($)":
            new_df.insert(i, col, df[col].fillna(0))
            new_df[col] = new_df[col].astype(np.float64)
        elif col == "Invoice Date" or col == "Date Paid":
            df[col] = pd.to_datetime(df[col])
            df[col] = df[col].dt.strftime('%m/%d/%Y')
            new_df.insert(i, col, df[col])
        else:
            new_df.insert(i, col, df[col].fillna(''))
            new_df[col] = new_df[col].astype('string')
    
    if 'Disputed' in df:
        new_df['Disputed'] = df['Disputed'].astype('string')
    else:
        new_df.insert(len(df.columns), 'Disputed', '')
    return new_df
    
def get_create_dispute_page(driver: WebDriver):
    driver.get("https://retaillink2.wal-mart.com/apdp/CreateDisputes")
    return

def signal_handler(sig, frame):
    logging.info('You pressed Ctrl+C!')
    sys.exit(0)

def catch_up(df: pd.DataFrame) -> int:
    for index in range(len(df)):
        # if a claim's deduction code is 94 AND amount paid is < 0 AND hasn't already been disputed, then file a dispute
        dispute_criteria = (df.at[index, 'DEDUCTION CODE'] == "MERCHANDISE RETURN - DEFECTIVE MERCHANDISE [0094]" and df.at[index, 'Amount Paid($)'] < 0 and (pd.isnull(df.at[index, 'Disputed'])))
        if dispute_criteria:
            return index

main()