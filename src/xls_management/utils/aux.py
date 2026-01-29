

def all_in_sequence(sequence_0:list|tuple, sequence_1:list|tuple) -> bool:
    """
    Returns True if each value in sequence_0 is in sequence_1; elsewhere False
    """
    for value in sequence_0:
        if value not in sequence_1:
            return False
    return True