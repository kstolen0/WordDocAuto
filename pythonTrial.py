from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import os
import json

dir_path = os.path.dirname(os.path.realpath(__file__)) 
template = dir_path + "\\TrialDoc.docx"

doc = MailMerge(template)
vals = doc.get_merge_fields()

print(vals)

doc.merge(
    doc_name = 'hello world',
    meeting_time = '11'
    )

items = [{
        'field_id' : '3.1',
        'field_name' : 'testing',
        'field_time' : '1100'
    },{
        'field_id' : '3.2',
        'field_name' : 'another',
        'field_time' : '1110'
        }]
doc.merge_rows('field_id',items)
doc.write('newFile.docx')

'''
kwargs = {}

for n in vals:
    kwargs[n] = input("give me the " + n + ":  ")

doc.merge(**kwargs)
doc.write('newFile.docx')
'''
