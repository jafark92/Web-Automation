import pandas as pd
from time import sleep
from os import listdir
from faker import Faker
from shutil import rmtree
from selenium import webdriver
from tempfile import gettempdir
from openpyxl import load_workbook
from random import randint, choice
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support import expected_conditions as EC

def delete_temp_file():
    files = listdir(gettempdir())
    for f in files:
        try:
            rmtree(f)
        except:
            pass

def update_file(updated_df, sheet_name, main_file):    
    xls = pd.ExcelFile(main_file)
    original_df = pd.read_excel(xls, sheet_name)
    writer=pd.ExcelWriter(main_file)
    original_df.to_excel(writer, sheet_name=sheet_name)

    #read the existing sheets so that openpyxl won't create a new one later
    book = load_workbook(main_file)
    writer = pd.ExcelWriter(main_file, engine='openpyxl') 
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    updated_df.to_excel(writer, sheet_name, index=False)

    writer.save()



def site_rwgmobile(driver):
    file_name = "SIMS to order.xlsx"
    sheet_name = '1. RWG MOBILE'
    
    df1 = pd.read_excel(file_name, sheet_name, keep_default_na="")
    postcode = df1['POSTCODE']
    add1 = df1['ADDRESS LINE 1']
    add2 = df1['ADDRESS LINE 2 (Optional)']
    town = df1['TOWN']
    status = df1["Status"]
    
    try: # try if there is no pending row
        row = status[status=='pending'].index[0]
    
        # generate random 10 digits number. 0 is not consider in start
        mobile_number = "".join([str(randint(1,9))]+[str(randint(0,9)) for _ in range(9)])

        url = "https://rwgmobile.wales/order-free-rwg-mobile-sim-card/"
        for i in [row]:
            name = fake.name()
            email = name.replace(" ", "")+"@gmail.com"            
            try:
                driver.get(url)

                # Handle Cookies
                try:
                    driver.find_element_by_id("pea_cook_btn").click()
                    sleep(2)
                except:
                    pass

                # fill form
                driver.find_element_by_css_selector(".your-name > input:nth-child(1)").send_keys(name)
                driver.find_element_by_css_selector(".wpcf7-tel").send_keys(mobile_number)
                driver.find_element_by_css_selector(".wpcf7-email").send_keys(email)
                driver.find_element_by_id("i-first_line").send_keys(add1[i])
                driver.find_element_by_id("i-second_line").send_keys(add2[i])
                driver.find_element_by_id("i-post_town").send_keys(town[i])
                driver.find_element_by_id("postcode-lookup").send_keys(postcode[i])
                sleep(2)

                # Click on Submit buttonw2
                driver.find_element_by_css_selector(".wpcf7-submit").click()
                sleep(2)

                # Delete cookies & Temp Files
                driver.delete_all_cookies()
                delete_temp_file()

                # update status in original file
                df1 = df1.copy()
                df1.Status.iloc[row] = "ordered"
                update_file(df1, sheet_name, file_name)

                print("Soccessfuly ordered")
            except:
                print("failed to Ordered")
    except IndexError:
        print("There is no pending order in", sheet_name)


def site_lebara(driver):
    file_name = "SIMS to order.xlsx"
    sheet_name = '4. Lebara'
    
    df1 = pd.read_excel(file_name, sheet_name, keep_default_na="")
    postcode = df1['POSTCODE']
    add1 = df1['ADDRESS LINE 1']
    add2 = df1['ADDRESS LINE 2 (Optional)']
    city = df1['CITY']
    delivery_info = df1['+ Additional Delivery Info']
    status = df1["Status"]
    
    try:
        row = status[status=='pending'].index[0]

        try:
            url = 'https://mobile.lebara.com/gb/en/free-sim'
            for i in [row]:
                name = fake.name()
                email = name.replace(" ", "")+"@gmail.com"
                first_name = name.split(' ')[0].replace(".", "")
                last_name = name.split(' ')[1].replace(".", "")

                driver.get(url)

                try:
                    driver.find_element_by_xpath(
                        "//*[@id='cookiesConsentModal']/div/div/div[2]/button[2]").click()
                except:
                    pass


                sleep(5)
                driver.find_element_by_id("buyNowFormSIM").click()

                # click on enter address manually
                WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/main/div[6]/div/div/div[3]/div/div/div/form/div[5]/div[2]/div[2]/a[2]"))
                ).click()

                #click  on Additional info
                driver.find_element_by_xpath("/html/body/main/div[6]/div/div/div[3]/div/div/div/form/div[5]/div[2]/div[4]/div/div/div[1]").click()

                #start filling form
                driver.find_element_by_id("register.firstName").send_keys(first_name)
                driver.find_element_by_id("register.lastName").send_keys(last_name)
                driver.find_element_by_id("register.email").send_keys(email)
                driver.find_element_by_id("addressLine1").send_keys(add1[i])
                driver.find_element_by_id("addressLine2").send_keys(add2[i])
                driver.find_element_by_id("postCode").send_keys(postcode[i])
                driver.find_element_by_id("city").send_keys(city[i])
                driver.find_element_by_id("additionalInfo").send_keys(delivery_info[i])
                sleep(5)

                # Click on Submit Button
                driver.find_element_by_id("submitOrderBtn").submit()
                sleep(5)

                # Delete Cppkies and Temp files
                driver.delete_all_cookies()
                delete_temp_file()

                # update status in original file
                df1 = df1.copy()
                df1.Status.iloc[row] = "ordered"
                update_file(df1, sheet_name, file_name)

                print("Successfully Ordered")    
        except:
            print("Failed to Ordered")
    except IndexError:
        print("There is no pending order in", sheet_name)


def site_vectone(driver):
    file_name = "SIMS to order.xlsx"
    sheet_name = '5. Vectone'
    
    df1 = pd.read_excel(file_name, sheet_name, keep_default_na="")
    postcode = df1['POSTCODE']
    add1 = df1['ADDRESS LINE 1']
    start_address = df1['Search for Starting Address']
    phone_number = df1['PHONE NUMBER']
    status = df1["Status"]
    
    try:
        row = status[status=='pending'].index[0]
    
        # generate random phone number if it is empty
        phone_number = choice(["078", "079", "077", "+4478", "+4479", "+4477"])+ "".join([str(randint(0,8)) for _ in range(8)]) if phone_number[row]=="" else phone_number[row]
    
        try:
            url = 'https://www.vectonemobile.co.uk/vmfreesimorder/step2'
            driver.maximize_window()

            for i in [row]:
                name = fake.name()
                email = name.replace(" ", "")+"@gmail.com"
                first_name = name.split(' ')[0].replace(".", "")
                last_name = name.split(' ')[1].replace(".", "")

                driver.get(url)

                driver.execute_script('document.getElementById("edit-no-of-sims").value=3')
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight*%s);" % .1)
                driver.find_element_by_xpath("//*[contains(@class,'col-12 col-sm-12 col-md-6 col-lg-6 summary-next-btn float-right')]").click()

                #  fill form
                WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.ID, "first_name"))
                ).send_keys(first_name)
                driver.find_element_by_id("last_name").send_keys(last_name)
                driver.find_element_by_id("email").send_keys(email)
                driver.find_element_by_id("contact_number").send_keys(phone_number)
                driver.find_element_by_id("opc_input").send_keys(postcode[i])
                driver.find_element_by_id("opc_button").click()

                # search for address
                try:
                    sleep(3)
                    driver.find_element_by_xpath('//*[@id="opc_dropdown"]//option[contains(text(),"'+start_address[i]+'")]').click()
                except:
                    print("Could not able to find with that Starting Address(Make sure it is case sensitive)")
                    pass

                # click submit button
                driver.find_element_by_xpath("//*[contains(@class,'col-sm-12 col-md-12 col-12 summary-next-btn text-right')]").click()
                sleep(5)

                # Delete Cookies and temp files
                driver.delete_all_cookies()
                delete_temp_file()

                # update status in original file
                df1 = df1.copy()
                df1.Status.iloc[row] = "ordered"
                update_file(df1, sheet_name, file_name)

            print("Successfully Ordered")
        except:
            print("failed to order")
    except IndexError:
        print("There is no pending order in", sheet_name)


def site_vodafone(driver):
    file_name = "SIMS to order.xlsx"
    sheet_name = '6. Vodafone'
    
    df1 = pd.read_excel(file_name, sheet_name, keep_default_na="")
    add1 = df1['ADDRESS LINE 1']
    add2 = df1['ADDRESS LINE 2']
    postcode = df1['POSTCODE']
    town = df1['TOWN/CITY']
    status = df1["Status"]

    try:
        row = status[status=='pending'].index[0]
    
        try:
            #url = 'https://freesim.vodafone.co.uk/check-out-payg'
            url = "https://v3.lolagrove.com/LeadPages/Vodafone.109/Vodafone.224/VodafoneFormTesting.5351/new_design/payg-summersim.aspx?id=24165.6850&get_referrer=&"
            driver.maximize_window()
            for i in [row]:
                name = fake.name()
                email = name.replace(" ", "")+"@gmail.com"
                first_name = name.split(' ')[0].replace(".", "")
                last_name = name.split(' ')[1].replace(".", "")
                print("start geting url")
                driver.get(url)
                print("success in geting url")

                try:
                    driver.find_element_by_xpath(
                        '//*[contains(@class,"optanon-alert-box-button-middle accept-cookie-container")]').click()
                    sleep(2)
                except:
                    pass

                driver.find_element_by_id("blank-sim").click()\

                WebDriverWait(driver,5).until(
                    EC.presence_of_element_located((By.ID, "txtFirstName"))
                ).send_keys(first_name)
                driver.find_element_by_id("txtLastName").send_keys(last_name)
                driver.find_element_by_id("txtEmail").send_keys(email)
                driver.find_element_by_id("txtPostCodeLookup").send_keys(postcode[i])
                driver.find_element_by_id("l-trigger-find-address").click()

                WebDriverWait(driver.find_element_by_id("addressLookup"),5).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, "option"))
                )[1].click()

                driver.find_element_by_id("txtAddress1").clear()
                driver.find_element_by_id("txtAddress1").send_keys(add1[i])

                driver.find_element_by_id("txtAddress2").clear()
                driver.find_element_by_id("txtAddress2").send_keys(add2[i])

                driver.find_element_by_id("txtTownCity").clear()
                driver.find_element_by_id("txtTownCity").send_keys(town[i])

                driver.find_element_by_id("chkPrivacy").click()
                sleep(2)

                # Click on Si=ubmit button
                driver.find_element_by_id("ibSubmit").click()                

                # delete cookies and temp files
                driver.delete_all_cookies()
                delete_temp_file()

                # update status in original file
                df1 = df1.copy()
                df1.Status.iloc[row] = "ordered"
                update_file(df1, sheet_name, file_name)

            print("Successfully Ordered")
        except:
            print("Failed to order")
    except IndexError:
        print("There is no pending order in", sheet_name)

# Main Program

driver = webdriver.Firefox(executable_path=GeckoDriverManager().install())

fake = Faker()

site_rwgmobile(driver)

sleep(randint(10,15)) # wait for 10 to 15 seconds before moving to next site
site_lebara(driver)

sleep(randint(10,15))
site_vectone(driver)

sleep(randint(10,15))
site_vodafone(driver)

driver.quit()
