import os

import streamlit as st
from backend import fill_xl_student, read_workbook, FOLDER
from config import *
from password import *
import shutil

st.title('Л2Ш Заказ одежды')

first_name = st.text_input('Имя ребенка').replace(' ', '')
second_name = st.text_input('Фамилия ребенка').replace(' ', '')

col1, col2 = st.columns(2)

with col1:
    group_num = st.selectbox('Класс:', range(6, 12))
with col2:
    group_letter = st.selectbox('Буква', ["А", "Б", "В", "Г", "Д", "Е"])

result = {'group': str(group_num) + group_letter, 'first_name': first_name, 'second_name': second_name}
positions = []
for cloth in clothes:
    title, price = cloth.rsplit(' ', 1)[0], int(cloth.split()[-1])
    st.subheader(f"{title} {price} руб")
    col4, col5 = st.columns(2)
    position = {'title': title, 'price': price, 'size': NOT_CHOSEN}
    with col4:
        if title in no_size_titles:
            position['size'] = 'NO SIZE'
        else:
            position['size'] = st.selectbox('Выберите размер:', sizes, key=cloth)
    with col5:
        position['number'] = st.number_input('Укажите кол-во', min_value=0, max_value=100, key=cloth + 'num')
    st.text('\n\n\n\n')
    positions.append(position)
result['positions'] = positions


if len(first_name) > 0 and len(second_name) > 0:
    pwd = st.text_input('Введите пароль: ', type='password')
    login_success = st.button('Отправить форму')
    if login_success:
        if check_user_password(pwd):
            fill_xl_student(result)
            st.text('Пожалуйста, скачайте файл и проверьте, что Ваши данные записались верно')
            st.download_button(label=f'Скачать Excel таблицу ({second_name})',
                               data=read_workbook(result['group']),
                               file_name=f"{result['group']}.xlsx",
                               mime='application/octet-stream')
        else:
            st.toast('Неверный пароль!')
else:
    st.text('Пожалуйста, заполните поля "Имя" и "Фамилия"')


if st.checkbox('Я администратор'):
    admin_pass = st.text_input('Введите пароль: ', type='password', key='admin')
    if check_admin_password(admin_pass):
        shutil.make_archive('data', 'zip', os.path.join(os.getcwd(), 'data'))
        with open("data.zip", "rb") as fp:
            st.download_button(label=f'Скачать архив со всеми файлами',
                               data=fp,
                               file_name='data.zip',
                               mime="application/zip")



