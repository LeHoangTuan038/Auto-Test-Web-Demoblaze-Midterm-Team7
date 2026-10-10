import pytest
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from helpers import (
    PASSWORD,
    accept_alert,
    create_username,
    logged_in_username,
    log_in,
    open_auth_modal,
    sign_up,
    wait_for,
)


@pytest.mark.smoke
def test_sign_up_and_login_successfully(browser):
    username = create_username()
    assert sign_up(browser, username) == "Sign up successful."

    log_in(browser, username)

    assert username in logged_in_username(browser, username)


@pytest.mark.smoke
def test_user_can_log_out(browser):
    username = create_username()
    assert sign_up(browser, username) == "Sign up successful."
    log_in(browser, username)
    logged_in_username(browser, username)

    browser.find_element(By.ID, "logout2").click()

    wait_for(browser, EC.visibility_of_element_located((By.ID, "login2")))
    assert browser.find_element(By.ID, "signin2").is_displayed()


@pytest.mark.regression
def test_login_rejects_wrong_password(browser):
    username = create_username()
    assert sign_up(browser, username) == "Sign up successful."
    log_in(browser, username, "Wrong-pass-456")

    assert accept_alert(browser) == "Wrong password."


@pytest.mark.regression
def test_login_requires_username(browser):
    open_auth_modal(browser, "logInModal", "login2")
    browser.find_element(By.ID, "loginpassword").send_keys(PASSWORD)
    browser.find_element(
        By.CSS_SELECTOR, "#logInModal button[onclick='logIn()']"
    ).click()

    assert accept_alert(browser) == "Please fill out Username and Password."


@pytest.mark.regression
def test_login_requires_password(browser):
    open_auth_modal(browser, "logInModal", "login2")
    browser.find_element(By.ID, "loginusername").send_keys(create_username())
    browser.find_element(
        By.CSS_SELECTOR, "#logInModal button[onclick='logIn()']"
    ).click()

    assert accept_alert(browser) == "Please fill out Username and Password."


@pytest.mark.regression
def test_sign_up_rejects_duplicate_username(browser):
    username = create_username()
    assert sign_up(browser, username) == "Sign up successful."

    assert sign_up(browser, username) == "This user already exist."
