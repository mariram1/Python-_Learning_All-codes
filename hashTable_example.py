def find_pairs(arr, target):
    hash_table = {}
    pairs = []
    for num in arr:
        complement = target - num  # What we need to hit target
        if complement in hash_table:  # O(1) lookup
            pairs.append((complement, num))
        hash_table[num] = True  # Store num as a key
    return pairs

arr = [3, 5, 7, 2, 8]
print(find_pairs(arr, 10))  # Output: [(3, 7), (2, 8)]
