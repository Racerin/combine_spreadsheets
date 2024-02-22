import re

# FUNCTIONS
def ensure_str_unique(input_str:str, str_list:list[str]=[]):
    """ 
    If 'input_str' is inside of list, update it's name and try again to add it.
    Return the accepted str.
    'str_list' automatically updates with each tested 'input_str'
    """
    n_str_list = len(str_list)
    for _ in range(n_str_list + 1):
        if input_str in str_list:
            # Append to 'input_str' name and try again
            input_str = input_str + "X"
        else:
            str_list.append(input_str)
            break
    else:
        raise OverflowError("Code Error: Str options exceeded.")
    return input_str

def is_excel_file(file_name):
    return any((file_name.endswith(ext) for ext in ('.xls','.xlsx')))

def reduce_whitespace_to_one_space(str1:str):
    # return str1
    return re.sub(r'\s+'," ", str1).strip()

def dict__old_to_unique_new_field(
        columns:list[str], 
        strip:bool=False,
        titleize:bool=False,
        lower:bool=False,
        upper:bool=False,
        replace_original_sub_str:list[tuple()]=[],
        )->dict(str=str):
    # reduce whitespace to one space
    new_strs = [re.sub(r'\s+'," ",old_str) for old_str in columns]
    # Apply to strings according to flag
    if strip: new_strs = [str1.strip() for str1 in new_strs]
    if titleize: new_strs = [str1.title() for str1 in new_strs]
    if lower: new_strs = [str1.lower() for str1 in new_strs]
    if upper: new_strs = [str1.upper() for str1 in new_strs]
    if replace_original_sub_str:
        for original,to_replace in replace_original_sub_str:
            new_strs = [str1.replace(original,to_replace) for str1 in new_strs]
    assert all((not s.endswith(" ") and not s.startswith(" ") for s in new_strs))
    # Make each unique str
    uniques = list()
    for str1 in new_strs:
        n_uniques = len(uniques)
        for _ in range(n_uniques+1):
            if str1 in uniques:
                # Rename str1 and try again
                str1 += "X"
            else:
                # Go ahead and add the str1
                break
        else:
            raise OverflowError(f"'{str1}' should be found in list '{new_strs}'.")
        uniques.append(str1)
    assert len(uniques) == len(columns), str(uniques)
    # convert to dict with old_str:new_str format
    ret_dict = {old:new for (old,new) in zip(columns,uniques)}
    return ret_dict