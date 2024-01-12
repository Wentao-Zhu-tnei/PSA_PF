def print_spacer():
    return print('-'*65)

def print_double_spacer():
    return print('='*65)

def print_list(list):
    list_len=len(list)
    print_spacer()
    print(f'There are {list_len} elements in the list')
    return print(*list, sep='\n')

def print_several_lists(*args):

    list_of_lists=[]
    for idx, arg in enumerate(args):
        list_of_lists.append(arg) 
    print_double_spacer()
    for list in zip(*list_of_lists):
        print(*list, sep=' |--| ')

    return None


def print_list_numbered(list):

    list_num=[]
    orig_list=[]

    print_double_spacer()
    for idx,elem in enumerate(list):
        list_num.append(idx)        
        orig_list.append(elem)

    return print_several_lists([list_num,orig_list])