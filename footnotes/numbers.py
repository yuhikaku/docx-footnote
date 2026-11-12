def find_missing_numbers(numbers: list[int]) -> list[int]:
    """連続する整数列（min~max）に対する欠番を返す"""
    if not numbers:
        return []
    number_set = set(numbers)
    return [n for n in range(min(numbers), max(numbers) + 1) if n not in number_set]
