def find_unused_numbers(max_num: int, used: set[int]) -> list[int]:
    """
    max_num までの整数のうち、used に含まれないものを返す
    例: max_num=5, used=[1,3] の場合、[2,4,5] を返す
    例: max_num=5, used=[1,2,3,4,5] の場合、[] を返す
    例: max_num=5, used=[] の場合、[1,2,3,4,5] を返す
    """
    all_range = range(1, max_num + 1)
    return [n for n in all_range if n not in used]
