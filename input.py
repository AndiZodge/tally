from fastapi import FastAPI, Response
from io import StringIO
from fastapi.responses import StreamingResponse
import pandas as pd
import xml.etree.ElementTree as ET

from pydantic import BaseModel
app = FastAPI()

class TallyData(BaseModel):
    REFERENCEDATE: str
    NARRATION: str
    VOUCHERTYPENAME: str
    PARTYLEDGERNAME: str 
    COST: float
    VOUCHERNUMBER: str
    
class Prase(BaseModel):
    VOUCHER: list[TallyData]
    
@app.get("/")
def download_excel():
    tree = ET.parse('input.xml')
    root_item = tree.getroot()
    tally_list = []
    for tag in root_item.findall('.//VOUCHER'):
        for tag2 in tag.findall('./ALLLEDGERENTRIES.LIST'):
            total_temp_amt = 0
            temp_amt = tag2.find('AMOUNT')
            if temp_amt != None:
                total_temp_amt += float(temp_amt.text)

        v_name = tag.find('VOUCHERTYPENAME')
        if v_name != None:
            v_name = v_name.text
        else:
            v_name = "nan" 
        r_date = tag.find('REFERENCEDATE')
        if r_date != None:
            r_date = r_date.text
        else:
            r_date = "nan"
        nar = tag.find('NARRATION')
        if nar != None:
            nar = nar.text
        else:
            nar = "nan"
        p_name = tag.find('PARTYLEDGERNAME')
        if p_name != None:
            p_name = p_name.text
        else:
            p_name = "nan"
        v_no = tag.find('VOUCHERNUMBER')
        if v_no != None:
            v_no = v_no.text
        else:
            v_no = "nan"
        temp = TallyData(REFERENCEDATE=r_date,
                    NARRATION=nar,
                    VOUCHERTYPENAME=v_name,
                    PARTYLEDGERNAME=p_name,
                    COST=total_temp_amt,
                    VOUCHERNUMBER=v_no
                    )
        tally_list.append(temp)
    
    data = Prase(VOUCHER=tally_list)
    tally_dicts = [tally.dict() for tally in data.VOUCHER]

    output = StringIO()

    df = pd.DataFrame(tally_dicts)
    df.to_excel(output)
    file_name = "example.xlsx"

    response = StreamingResponse(iter([output.getvalue()]), media_type="text/xlsx")
    response.headers["Content-Disposition"] = f"attachment; filename={file_name}"
    
    return response
