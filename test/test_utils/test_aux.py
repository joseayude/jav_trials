from xls_management.utils.aux import all_in_sequence, list_from_comma_separated_str


def test_all_in_sequence():
    a3_0 = ('one', 'two', 'three')
    a3_1 = ('one', 'two', 'three')
    
    assert all_in_sequence(a3_0, a3_1) is True

    
    a2_0 = ('two', 'one')
    assert all_in_sequence(a2_0, a3_0) is True
    assert all_in_sequence(a3_0, a2_0) is False
    a0:tuple = ()
    
    assert all_in_sequence(a0, a2_0) is True
    assert all_in_sequence(a3_0, a0) is False

def test_get_comma_separated_list():   
    assert list_from_comma_separated_str('one, two, three') == ['one', 'two', 'three']
    assert list_from_comma_separated_str(' one , two , three ') == ['one', 'two', 'three']
    assert list_from_comma_separated_str('one,?two,three') == ['one', 'two', 'three']
    assert list_from_comma_separated_str('one,\t?two ,th ree\n') == ['one', 'two', 'three']
    assert list_from_comma_separated_str('one,\t\t\t\t?two ,th ree\n') == ['one', 'two', 'three']
    #it works as well as a string method.
    assert 'one,\t\t\t\t?two ,th ree\n, '.get_comma_separated_list() == ['one', 'two', 'three']
    