'''surveyexport module testing'''


import os
from app import go
import json


def test_go():
    '''testing go function from surveyexport module with a big survey in json format'''
    for f in os.listdir("./events"):
        data = open("./events/" + f)
        inp = json.load(data)
        go(inp["body"])
        assert "survey_export.docx" in os.listdir("/tmp")
        os.remove("/tmp/survey_export.docx")

test_go()
