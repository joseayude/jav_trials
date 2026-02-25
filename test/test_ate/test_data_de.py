
from itertools import islice


def test_TDVKAttribute():
    from xls_management.ate.data_de import TDVCAttribute as VC

    assert VC.ID == 'ID'
    assert VC.RequirementBased == 'Basierend auf der Anforderung'
    assert VC.Status == 'Status'
    assert VC.Temp1Text == 'Temp1_text'
    assert VC.Action == 'Aktion'

    all_attributes = [attribute.value for attribute in VC]
    assert len(all_attributes) == 5
    assert all_attributes[0] == 'ID'
    assert all_attributes[1] == 'Basierend auf der Anforderung'
    assert all_attributes[2] == 'Status'
    assert all_attributes[3] == 'Temp1_text'
    assert all_attributes[4] == 'Aktion'

    two_three_str = [attribute.value for attribute in islice(VC,1,len(VC)-1)]
    assert len(two_three_str) == 3
    assert 'ID' not in two_three_str
    assert two_three_str[0] == 'Basierend auf der Anforderung'
    assert two_three_str[1] == 'Status'
    assert two_three_str[2] == 'Temp1_text'
    assert 'Aktion' not in two_three_str

    my_index =  islice(VC,1,len(VC)-1)
    two_three = [attribute for attribute in my_index]
    assert len(two_three) == 3
    assert 'ID' not in two_three
    assert two_three[0] == VC.RequirementBased
    assert two_three[1] == VC.Status
    assert two_three[2] == VC.Temp1Text
    assert 'Aktion' not in two_three