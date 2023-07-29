import streamlit as st
from backend import fill_xl_student, read_workbook
from config import *
st.title('Л2Ш Заказ одежды')

first_name = st.text_input('Имя ребенка')
second_name = st.text_input('Фамилия ребенка')
third_name = st.text_input('Отчество ребенка')

col1, col2 = st.columns(2)

with col1:
    group_num = st.selectbox('Класс:', range(6, 12))
with col2:
    group_letter = st.selectbox('Буква', ["А", "Б", "В", "Г", "Д", "Е"])


result = {'group': str(group_num) + group_letter, 'first_name': first_name, 'second_name': second_name,
          'third_name': third_name}
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
if st.button('Отправить форму'):
    fill_xl_student(result)


st.download_button(label=f'Скачать Excel таблицу ({second_name})',
                   data=read_workbook(result['group']),
                   file_name=f"{result['group']}.xlsx",
                   mime='application/octet-stream')
