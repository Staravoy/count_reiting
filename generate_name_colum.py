def generate_name_column():
    first_char = ord('A')
    last_char = ord('Z')


    list_name_column = []
    list_index_column = []
    for x in range(first_char, last_char + 1):
        list_name_column.append(chr(x))

    for char1 in range(first_char, last_char + 1):
        for char2 in range(first_char, last_char + 1):
            two_char = chr(char1) + chr(char2)
            list_name_column.append(two_char)

    for index_column in range(1, len(list_name_column)):
        list_index_column.append(index_column)

    make_dic = dict(zip(list_name_column, list_index_column))
    return make_dic