import http.client
import json
import os.path
import urllib.parse
import openpyxl

TARGET_DIRECTORY = r'confidential'
XL_FILE = 'confidential'


def create_mom_objects(row):
    mom_object = make_mom(row)
    return mom_object


def make_mom(row):
    mom = MOM(row)
    return mom


class MOM(object):
    def _init_(self, row):
        self.name = row[0].value
        self.year = row[2].value.year
        self.month = row[2].value.month
        self.year_month = f"{self.year}-{self.month}"
        self.station = row[1].value
        self.station = self.station[:4]
        if self.station == "Gene":
            self.station = "General"


def make_folder(mom_object):
    try:
        os.makedirs(
            fr"""{TARGET_DIRECTORY}\{mom_object.year}\{mom_object.year_month}\{mom_object.station}\{mom_object.name.replace('"', "").replace(' / ', ' - ').replace(':', '-').replace('#', '')}""")
    except FileExistsError:
        return fr"""{TARGET_DIRECTORY}\{mom_object.year}\{mom_object.year_month}\{mom_object.station}\{mom_object.name.replace('"', "").replace(' / ', ' - ').replace(':', '-').replace('#', '')}"""
    else:
        return fr"""{TARGET_DIRECTORY}\{mom_object.year}\{mom_object.year_month}\{mom_object.station}\{mom_object.name.replace('"', "").replace(' / ', ' - ').replace(':', '-').replace('#', '')}"""


def postman_search(mom_object):
    conn = http.client.HTTPSConnection("ibi.mfs.cloud")
    payload = ''
    headers = {
        'X-Authentication': '',
        'Content-Type': 'application/json',
        'Cookie': 'ASP.NET_SessionId=; fileDownload=true',
    }
    query = mom_object.name
    search = urllib.parse.quote(query)
    conn.request("GET", f"/REST/objects?p0={search}",
                 payload, headers)
    res = conn.getresponse()
    data = res.read()
    data_decoded = json.loads(data.decode("utf-8"))
    try:
        mom_object.files = data_decoded["Items"][0]["Files"]
    except IndexError:
        errors.append(f'{mom_object.name} - no files in mfiles folder')
    except KeyError:
        errors.append(f'{mom_object.name} - no search results in mfiles')
    else:
        results_list = []

        for i in range(len(mom_object.files)):
            file_id = 0
            if mom_object.files[i]["Extension"].lower() == 'pdf':
                file_id = mom_object.files[i]["ID"]
                file_name = f'{mom_object.files[i]["Name"]}.pdf'
                if file_id != 0:
                    dict_entry = {
                        "object_type": data_decoded["Items"][0]["ObjVer"]["Type"],
                        "object_id": data_decoded["Items"][0]["ObjVer"]["ID"],
                        "object_version": data_decoded["Items"][0]["ObjVer"]["Version"],
                        "file_id": file_id,
                        "file_name": file_name,
                    }
                    results_list.append(dict_entry)
                else:
                    dict_entry = {
                        "object_type": 0,
                        "object_id": 0,
                        "object_version": 0,
                        "file_id": 0,
                        "file_name": 0,
                    }
                    results_list.append(dict_entry)
        return results_list


def download(mom_object):
    conn = http.client.HTTPSConnection("mfs.cloud")
    payload = ''
    headers = {
        'X-Authentication': '',
        'Content-Type': 'application/json',
        'Cookie': 'ASP.NET_SessionId=; fileDownload=true'
    }
    if mom_object.postman:
        mom_object.file_count = len(mom_object.postman)
        if mom_object.file_count == 0:
            errors.append(errors.append(f'{mom_object.name}'))
        for i in range(mom_object.file_count):
            try:
                request = f"/REST/objects/{mom_object.postman[i]['object_type']}/{mom_object.postman[i]['object_id']}/" \
                          f"{mom_object.postman[i]['object_version']}/files/{mom_object.postman[i]['file_id']}/content"
            except TypeError:
                errors.append(f'{mom_object.name}')
            else:
                conn.request("GET", request, payload, headers)
                res = conn.getresponse()
                data = res.read()
                path = mom_object.folder.replace('"', "").replace(' / ', ' - ').replace('"', '')
                name = mom_object.postman[i]["file_name"]
                try:
                    file = open(os.path.join(path, name), 'wb')
                except FileNotFoundError:
                    errors.append(f'{mom_object.name}')
                except OSError:
                    right_chunk = path[3:]
                    right_chunk = right_chunk.replace(':', '-').replace('#', '')
                    path = path[:3] + right_chunk
                    name = name.replace(':', '-').replace('#', '').replace('"', '')
                    file = open(os.path.join(path, name), 'wb')
                    file.write(data)
                    file.close()
                else:
                    file.write(data)
                    file.close()


def master():
    global errors
    wrkbk = openpyxl.load_workbook(XL_FILE)
    sh = wrkbk.active
    errors = []
    start = 2
    row_num = start
    pdf_total = 0
    for row in sh.iter_rows(min_row=start, min_col=1, max_row=779, max_col=3):
        mom_object = create_mom_objects(row)
        mom_object.folder = make_folder(mom_object)
        mom_object.postman = postman_search(mom_object)
        mom_object.file_count = 0
        download(mom_object)
        pdf_total += mom_object.file_count
        print(f'row #: {row_num}')
        print(f'date: {mom_object.year_month}')
        print(f'name: {mom_object.name}')
        print(f'error count: {len(errors)}')
        print(f'total files: {pdf_total}')
        print(f'files added: {mom_object.file_count}\n')
        row_num += 1
        with open('errors.txt', 'w') as f:
            for i in range(len(errors)):
                try:
                    f.write(f'{i + 1}: {errors[i]}\n')
                except:
                    f.write(f'{i + 1}\n')


master()