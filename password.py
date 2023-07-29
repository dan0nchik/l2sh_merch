import os


def set_user_password(pwd):
    os.environ['USER_PASSWORD'] = pwd


def check_user_password(pwd):
    return os.environ['USER_PASSWORD'] == pwd


def set_admin_password(pwd):
    os.environ['ADMIN_PASSWORD'] = pwd


def check_admin_password(pwd):
    return os.environ['ADMIN_PASSWORD'] == pwd
