import os
import shutil

import streamlit as st

from backend import fill_xl_student, read_workbook, FOLDER, fill_xl_group
from config import *
from password import *
st.title('Л2Ш Заказ одежды')
second_name = st.text_input('Фамилия ребенка').replace(' ', '')
first_name = st.text_input('Имя ребенка').replace(' ', '')


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

st.header(':exclamation:Внимание:exclamation: Форму можно отправить только один раз:exclamation:')

if len(first_name) > 0 and len(second_name) > 0:
    pwd = st.text_input('Введите пароль: ', type='password')
    login_success = st.button('Отправить форму')
    if login_success:
        if check_user_password(pwd):
            fill_xl_student(result)
            fill_xl_group(result)
            st.subheader('Пожалуйста, скачайте файл и проверьте, что Ваш заказ правильно записался в общую таблицу')
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
        st.text(os.listdir(os.getcwd()))

        file_name = st.selectbox('Выберите год: ', [str(i) for i in range(2023, 2030)])
        if st.button('Сформировать архив'):
            shutil.make_archive(file_name, 'zip', os.path.join(os.getcwd(), file_name))
            with open(f"{file_name}.zip", "rb") as fp:
                st.download_button(label=f'Скачать',
                                   data=fp,
                                   file_name=f'{file_name}.zip',
                                   mime="application/zip")
        file = st.selectbox('Выберите файл: ', os.listdir(os.getcwd()))
        if st.button('Удалить'):
            if os.path.isdir(file):
                shutil.rmtree(file)
            os.remove(file)
            st.toast(f'Файл {file} удален!')