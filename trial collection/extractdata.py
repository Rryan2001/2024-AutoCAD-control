import ezdxf
import json
import win32com.client as win32
# 矩形对角顶点
rectangle = ((0, 0), (100, 100))

#def extract_text_data(filename):
def extract_text_data():
    #doc = ezdxf.readfile(filename)
    doc = wincad.ActiveDocument
    msp = doc.modelspace()

    text_data = []
    for text in msp.query("TEXT"):
        x, y, _ = text.dxf.insert
        if rectangle[0][0] <= x <= rectangle[1][0] and rectangle[0][1] <= y <= rectangle[1][1]:
            data = {}

            data['text'] = text.dxf.text
            data['location'] = (text.dxf.insert[0], text.dxf.insert[1])
            data['rotation'] = text.dxf.rotation
            data['color'] = text.dxf.color

            text_data.append(data)
    for mtext in msp.query("MTEXT"):
        x, y, _ = mtext.dxf.insert
        if rectangle[0][0] <= x <= rectangle[1][0] and rectangle[0][1] <= y <= rectangle[1][1]:
            data = {}

            data['text'] = mtext.plain_text()
            data['location'] = (mtext.dxf.insert[0], mtext.dxf.insert[1])
            data['rotation'] = mtext.dxf.rotation
            data['color'] = mtext.dxf.color

            text_data.append(data)

    return text_data

def output_json(filename, data):
    with open(filename, 'w', encoding='utf-8') as file:
        json.dump(data, file, indent=4, ensure_ascii=False)



if __name__ == "__main__":
    dxf_filename = "*.dxf"
    json_filename = "*.json"

    text_data = extract_text_data(dxf_filename)
    output_json(json_filename, text_data)


