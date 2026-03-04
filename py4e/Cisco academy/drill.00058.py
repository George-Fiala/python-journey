def get_longest_word(text):
    if not text:
        return None
    
    words = text.split()
    longest = None

    
    for word in words:
        if longest is None or len(word) > len(longest):
            longest = word
    return longest

