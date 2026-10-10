from uuid import uuid4

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


BASE_URL = "https://demoblaze.com/"
PASSWORD = "Test-pass-123"
WAIT_SECONDS = 15


def wait_for(driver, condition, timeout=WAIT_SECONDS):
    return WebDriverWait(driver, timeout).until(condition)


def create_username():
    return f"qa_{uuid4().hex[:12]}"


def open_home(driver):
    driver.get(BASE_URL)
    wait_for(driver, EC.visibility_of_element_located((By.CSS_SELECTOR, "#tbodyid .card")))


def accept_alert(driver):
    wait_for(driver, EC.alert_is_present())
    alert = driver.switch_to.alert
    message = alert.text
    alert.accept()
    return message


def open_auth_modal(driver, modal_id, nav_id):
    open_home(driver)
    driver.find_element(By.ID, nav_id).click()
    wait_for(driver, EC.visibility_of_element_located((By.ID, modal_id)))


def sign_up(driver, username, password=PASSWORD):
    open_auth_modal(driver, "signInModal", "signin2")
    driver.find_element(By.ID, "sign-username").send_keys(username)
    driver.find_element(By.ID, "sign-password").send_keys(password)
    driver.find_element(
        By.CSS_SELECTOR, "#signInModal button[onclick='register()']"
    ).click()
    return accept_alert(driver)


def log_in(driver, username, password=PASSWORD):
    open_auth_modal(driver, "logInModal", "login2")
    driver.find_element(By.ID, "loginusername").send_keys(username)
    driver.find_element(By.ID, "loginpassword").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, "#logInModal button[onclick='logIn()']").click()


def logged_in_username(driver, username):
    welcome = wait_for(driver, EC.visibility_of_element_located((By.ID, "nameofuser")))
    wait_for(
        driver,
        lambda current_driver: username in current_driver.find_element(
            By.ID, "nameofuser"
        ).text,
    )
    return welcome.text


def add_product_by_index(driver, index=0):
    open_home(driver)
    products = wait_for(
        driver,
        lambda current_driver: current_driver.find_elements(
            By.CSS_SELECTOR, "#tbodyid .card h4 a"
        ),
    )
    products[index].click()
    wait_for(
        driver,
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#tbodyid .btn-success")),
    )
    driver.find_element(By.CSS_SELECTOR, "#tbodyid .btn-success").click()
    return accept_alert(driver)


def open_cart(driver, *, expect_items=False):
    driver.find_element(By.ID, "cartur").click()
    wait_for(driver, EC.url_contains("cart.html"))
    wait_for(driver, EC.presence_of_element_located((By.ID, "tbodyid")))
    if expect_items:
        wait_for(
            driver,
            lambda current_driver: current_driver.find_elements(
                By.CSS_SELECTOR, "#tbodyid tr"
            ),
        )


def cart_rows(driver):
    return driver.find_elements(By.CSS_SELECTOR, "#tbodyid tr")


def open_order_modal(driver):
    driver.find_element(By.XPATH, "//button[normalize-space()='Place Order']").click()
    wait_for(driver, EC.visibility_of_element_located((By.ID, "orderModal")))


def fill_order_form(driver, *, name="Test Buyer", card="4111111111111111"):
    fields = {
        "name": name,
        "country": "Vietnam",
        "city": "Hanoi",
        "card": card,
        "month": "10",
        "year": "2026",
    }
    for field_id, value in fields.items():
        driver.find_element(By.ID, field_id).send_keys(value)


def place_order(driver):
    driver.find_element(
        By.CSS_SELECTOR, "#orderModal button[onclick='purchaseOrder()']"
    ).click()
