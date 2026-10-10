import requests, sys, os

sys.stdout.reconfigure(encoding='utf-8')

files_to_upload = [
    (r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_المتوسط_الضبع_الاسود_LMS.xlsx', 'Medium_Hyena_Bank_LMS.xlsx'),
    (r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_النهائي_الضبع_الاسود_LMS.xlsx', 'Final_Hyena_Bank_LMS.xlsx')
]

print("=== Uploading to tmpfiles.org ===")
for local_path, upload_name in files_to_upload:
    try:
        with open(local_path, 'rb') as f:
            resp = requests.post('https://tmpfiles.org/api/v1/upload', files={'file': (upload_name, f)})
            data = resp.json()
            if data.get('status') == 'success':
                url = data['data']['url']
                # Direct download link: replace tmpfiles.org/ with tmpfiles.org/dl/
                dl_url = url.replace('tmpfiles.org/', 'tmpfiles.org/dl/')
                print(f"{upload_name} -> {dl_url}")
            else:
                print(f"Failed {upload_name}: {data}")
    except Exception as e:
        print(f"Error {upload_name}: {e}")

print("\n=== Uploading to filebin.net ===")
bin_id = "hyena-banks-2026-final"
for local_path, upload_name in files_to_upload:
    try:
        with open(local_path, 'rb') as f:
            headers = {'Filename': upload_name}
            resp = requests.post(f"https://filebin.net/{bin_id}/{upload_name}", data=f, headers=headers)
            if resp.status_code in [200, 201]:
                print(f"{upload_name} -> https://filebin.net/{bin_id}/{upload_name}")
            else:
                print(f"Failed filebin {upload_name}: {resp.status_code}")
    except Exception as e:
        print(f"Error filebin {upload_name}: {e}")
