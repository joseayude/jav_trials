

import re


def all_in_sequence(sequence_0:list|tuple, sequence_1:list|tuple) -> bool:
    """
    Returns True if each value in sequence_0 is in sequence_1; elsewhere False
    """
    for value in sequence_0:
        if value not in sequence_1:
            return False
    return True

#   Function EinlesenGetrennteWerteKomma(ByVal lahIDs As String) As Collection
    #Moved to utils as it could be used in multiple places.
def list_from_comma_separated_str(comma_separated_string:str) -> list:
    """
    Given commma separated string is cleaned and converted to list.
    The cleaning includes removing "?" and whitespace characters.
    
    :param comma_separated_string: A string with comma separated values, 
                which may include "?" characters and whitespace.
    :type comma_separated_string: str
    :return: list of cleaned values from the input string.
    :rtype: list
    """
    cleaned_string = re.sub('[\?\s]','',comma_separated_string)
    return cleaned_string.split(',')
#       Dim idCollection As Collection
#       Dim subStrings() As String
#       Set idCollection = New Collection
#       Dim newString As String
#       Dim x As Integer
#       
#       newString = Replace(lahIDs, "?", "")
#       subStrings = Split(newString, ",")
#       For x = LBound(subStrings) To UBound(subStrings)
#           idCollection.Add Trim(subStrings(x))
#       Next
#       Set EinlesenGetrennteWerteKomma = idCollection
#       End Function