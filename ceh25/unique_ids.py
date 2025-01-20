list_1 = [
    123
]

list_2 = [
    123,
    456
]

list_result = []

list_deleted = []

for i in list_2:
    if i not in list_1:
        list_result.append(i)
    else:
        list_deleted.append(i)

print(f'Первый список: {len(list_1)} элементов')
print(f'Удаленный список: {len(list_deleted)} элементов')

print(f'Второй список: {len(list_2)} элементов')
print(f'Итоговый список: {len(list_result)} элементов')

print()
print()
print(f'{list_result = }')
print(f'{list_deleted = }')
