import http.client
import json
import os.path
import urllib.parse
import openpyxl

XL_FILE = 'confidential.xlsx'
TARGET_DIRECTORY = r'D:\LTI-Related'


def create_mom_objects(row):
    mom_object = make_mom(row)
    return mom_object


def make_mom(row):
    mom = MOM(row)
    return mom


class MOM(object):
    def __init__(self, row):
        self.name = row[0].value
        self.year = row[1].value.year
        self.month = row[1].value.month
        self.year_month = f"{self.year}-{self.month}"
        self.station = row[2].value
        self.station = self.station[:4]
        if self.station == "Gene":
            self.station = "General"


def create_rel_objects(row):
    mom_object = make_mom(row)
    return mom_object


def make_rel(row):
    mom = MOM(row)
    return mom


def make_folder(mom_object):
    try:
        os.makedirs(
            fr"""{TARGET_DIRECTORY}\{mom_object.year}\{mom_object.year_month}\{mom_object.station}\{mom_object.name.replace('"', "").replace(' / ', ' - ').replace(':', '-').replace('#', '')[4:9]}""")
    except FileExistsError:
        return fr"""{TARGET_DIRECTORY}\{mom_object.year}\{mom_object.year_month}\{mom_object.station}\{mom_object.name.replace('"', "").replace(' / ', ' - ').replace(':', '-').replace('#', '')[4:9]}"""
    else:
        return fr"""{TARGET_DIRECTORY}\{mom_object.year}\{mom_object.year_month}\{mom_object.station}\{mom_object.name.replace('"', "").replace(' / ', ' - ').replace(':', '-').replace('#', '')[4:9]}"""


def file_search(mom_object):
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
        all_files_dict = {}
        file_list = []
        related_list = []

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
                    file_list.append(dict_entry)
                else:
                    dict_entry = {
                        "object_type": 0,
                        "object_id": 0,
                        "object_version": 0,
                        "file_id": 0,
                        "file_name": 0,
                    }
                    file_list.append(dict_entry)
                return dict_entry


def file_search2(related_dict, mom_object):  # dont need because i already have file id's
    conn = http.client.HTTPSConnection("ibi.mfs.cloud")
    payload = ''
    headers = {
        'X-Authentication': 'OHLvHwYoN5cdT4vyceVbqIK_MS8tc58FucuqHSiZs3Jo0-a-qJl2Hw6jy9YDtXxvZ4WsF_dq2Qi0LC25fMXc8Ojyu8UAz04Y8-eKInHEnp23L7xoB4lZpwFLuJ0y3VLVKoL-tJ4o5SFHVADVBHDmX-dYwN1-Ho6ulFE91ou3lgoaRtFgj1ZYaa7CmadgMurPMmxOSGQvOpZRFSnlOWzexVFnrAjV6hWHnlLDiN2Bwl3cuby5UnEgNBicUYMY3bKdmeWZkhpYWFIC_13PsIXi6pfqjb691xAjuuLa2FdAi6mgNMAlFwN0RXk-u9Bm9aE_FnIQCrpM7HRehheUVufHhQ',
        'Content-Type': 'application/json',
        'Cookie': 'ASP.NET_SessionId=tpypag5rw4momklcbvcduhdb; fileDownload=true',
    }
    for i in related_dict:
        query = i['Title']
        search = urllib.parse.quote(query)
        conn.request("GET", f"/REST/objects?p0={search}",
                     payload, headers)
        res = conn.getresponse()
        data = res.read()
        data_decoded = json.loads(data.decode("utf-8"))
        try:
            mom_object.related_files = data_decoded["Items"][0]["Files"]
        except IndexError:
            errors.append(f'{mom_object.name} - no files in mfiles folder')
        except KeyError:
            errors.append(f'{mom_object.name} - no search results in mfiles')
        else:
            all_files_dict = {}
            file_list = []
            related_list = []

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
                        file_list.append(dict_entry)
                    else:
                        dict_entry = {
                            "object_type": 0,
                            "object_id": 0,
                            "object_version": 0,
                            "file_id": 0,
                            "file_name": 0,
                        }
                        file_list.append(dict_entry)
                    return dict_entry


def get_related(dict_entry):
    conn = http.client.HTTPSConnection("")
    payload = ''
    headers = {
        'X-Authentication': '',
        'Content-Type': 'application/json',
        'Cookie': 'ASP.NET_SessionId=''; fileDownload=true',
    }
    try:
        conn.request("GET",
                     f"/REST/objects/{dict_entry['object_type']}/{dict_entry['object_id']}/"
                     f"{dict_entry['object_version']}/"
                     f"relationships?direction=both&type=objectversion",
                     payload, headers)
    except TypeError:
        pass
    else:
        res2 = conn.getresponse()
        data2 = res2.read()
        data_decoded2 = json.loads(data2.decode("utf-8"))
        return data_decoded2



def file_download(mom_object):
    conn = http.client.HTTPSConnection("ibi.mfs.cloud")
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


def related_download(related_dict, mom_object):
    conn = http.client.HTTPSConnection("ibi.mfs.cloud")
    payload = ''
    headers = {
        'X-Authentication': '',
        'Content-Type': 'application/json',
        'Cookie': 'ASP.NET_SessionId=; fileDownload=true'
    }
    if related_dict:
        folders = len(related_dict)
        for i in related_dict:
            filecount = len(i['Files'])
            m = 0
            for j in i['Files']:
                if j['Extension'].lower() == 'pdf':
                    print(f'{m}/{filecount - 1}/{folders}')
                    try:
                        request = f"/REST/objects/{i['ObjVer']['Type']}/{i['ObjVer']['ID']}/" \
                                  f"{i['ObjVer']['Version']}/files/{j['ID']}/content"
                    except TypeError:
                        errors.append(f'{mom_object.name}')
                    else:
                        conn.request("GET", request, payload, headers)
                        res = conn.getresponse()
                        data = res.read()
                        path = mom_object.folder.replace('"', "").replace(' / ', ' - ').replace('"', '')
                        name = j['Name'] + '.pdf'
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
                            mom_object.file_count += 1
                m += 1


def master():
    global errors
    wrkbk = openpyxl.load_workbook(XL_FILE)
    sh = wrkbk.active
    errors = []
    start = 1
    row_num = start
    pdf_total = 0
    for row in sh.iter_rows(min_row=start, min_col=1, max_row=3508, max_col=3):
        mom_object = create_mom_objects(row)
        mom_object.folder = make_folder(mom_object)
        mom_object.related = file_search(mom_object)
        related = get_related(mom_object.related)
        mom_object.file_count = 0
        related_download(related, mom_object)
        print(f'row #: {row_num}')
        print(f'date: {mom_object.year_month}')
        print(f'error count: {len(errors)}')
        print(f'files added: {mom_object.file_count}\n')
        row_num += 1
        with open('errors.txt', 'w') as f:
            for i in range(len(errors)):
                try:
                    f.write(f'{errors[i]}\n')
                except:
                    f.write(f'{i + 1}\n')


master()
