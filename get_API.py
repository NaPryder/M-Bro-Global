import requests
import json
import csv
import os
from Config import API_Config, PATH, extract_log

def download(CSV_URL, savefile, encoding='cp874', delimiter=','):
    with requests.Session() as s:
        download = s.get(CSV_URL)

        decoded_content = download.content.decode(encoding)
        cr = csv.reader(decoded_content.splitlines(), delimiter=delimiter)
        
        with open(savefile, 'w', encoding=encoding) as csvf:

            my_list = list(cr)
            csv_writer = csv.writer(csvf)
            
            csv_writer.writerows(my_list)
            extract_log(f"download file from api: {savefile}")

def decode_csv(csvfile_read, csvfile_write, encoding_from='cp874', decode_to='utf-8-sig'):
    temp =[]
    
    with open(csvfile_read, encoding=encoding_from) as myFile:  
        reader = csv.reader(myFile)
        for row in reader:
            temp.append(row)
    with open(csvfile_write, 'w', newline='\n', encoding=decode_to) as newfile:  
        writer = csv.writer(newfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for row in temp:
            if row:
                writer.writerow(row)

    extract_log(f"decode raw file save to: {csvfile_write}")

def clear_raw_data(raw_file, new_file):
    if os.path.exists(new_file):
        try:
            os.remove(raw_file)
            extract_log(f"delete raw file: {raw_file}")
        except:
            pass

def Main_request_api():
    Path = PATH()
    Api = API_Config()
    url = Api.request_url
    api_Key = Api.api_key
    year = Api.year

    Output_folder = Path.OUTPUT_PATH
    extract_log(F"--- Start request ---")
    extract_log(f" request url: {url}")
    extract_log(f" api-key: {api_Key}")
    extract_log(f" year: {year}")

    r = requests.get(url)
    d = r.json()
    extract_log(f"result json : ")
    extract_log(json.dumps(d, ensure_ascii=False , indent=4))

    download_url = d['result']['url'] + f"?api-key{api_Key}"
    extract_log(f" download url from ['result']['url'] : {download_url}")

    raw_file = os.path.join(Output_folder, f"raw_download_from_api_{year}.csv")
    new_file = os.path.join(Output_folder, f"download_from_api_{year}.csv")

    download(download_url, raw_file)
    decode_csv(raw_file , new_file)
    clear_raw_data(raw_file, new_file)


if __name__ == '__main__':
        
    Main_request_api()








