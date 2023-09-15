import os

import win32com.client

# Outlook 애플리케이션 열기
outlook = win32com.client.Dispatch("Outlook.Application")

# MAPI 네임스페이스 가져오기
namespace = outlook.GetNamespace("MAPI")

"""
outlook namespace 까지는 무조건 들어온다
"""
# print(namespace.GetDefaultFolder(6))

# wrong unknown
# print(namespace.Folders(1).Items.Folders.Item(1))

# 첫 번째 폴더 컬렉션 가져오기
folders = namespace.Folders.Item(1).Folders

"""
namespace에 Folders 중에 Folder는 하나밖에 없음
ex. 개인 메일 데이터 파일 또는 Outlook 데이터 파일
-> namespace.Folders.Item(1) = 데이터 파일
"""
for folder in namespace.Folders:
    print(folder)

"""
namespace.Folders('폴더 이름')하면 폴더 접근 가능
-> ex. Outlook 데이터 파일
-> but .Items로는 Item이 안보임
"""
# 환경변수 OUTLOOK_DATA_FILE 처리
OUTLOOK_DATA_FILE = os.environ.get("OUTLOOK_DATA_FILE")
folder_name = OUTLOOK_DATA_FILE
for folder in namespace.Folders().Items:
    print(folder.Name)

"""
폴더 추가는 이미 있는 폴더 이름이면 예외 발생
"""
folder_name = "add_folder_test"
new_folder = None
for folder in namespace.Folders.Item(1).Folders:
    if folder.Name == folder_name:
        new_folder = folder
        print(f"folder name 존재 : {folder_name}, 폴더 추가 실패")
        break
if new_folder is None:
    namespace.Folders.Item(1).Folders.Add(folder_name)
    print(f"폴더 추가 : {folder_name}")
"""
GetDefaultFolder로 DefaultFolder에 접근할 수 있음.
6번은 받은 편지함
"""
def print_items_in_folder(items) -> None:
    cnt = 0
    for item in items:
        if cnt == 3: break
        print(item)
        cnt += 1
inbox = namespace.GetDefaultFolder(6)
print(inbox) # 받은 편지함에 폴더 추가
print(inbox.Name) # 받은 편지함에 폴더 추가
print_items_in_folder(inbox.Items)
# inbox.Folders.Add("defaultfolder6 add test2") # 받은 편지함에 폴더 추가

"""
네임스페이스 내부 모든 폴더 출력
"""
for folder in namespace.Folders:
    print(folder)

"""

"""
for folder in namespace.Folders.Item(1).Folders:
    print(folder)

"""
10번은 내 연락처가 아니라 최상위 연락처

연락처 폴더 하위에 새 폴더 만들면 폴더가 두개 생겨버림

최상위 폴더 : 연락처
하위 폴더 : 추가하는 주소록 폴더(추가하면 하위에 있는 것 + 전체 주소록 폴더까지 두개 보이나??)

삭제하려면 하위로 가서 다 삭제해줘야함

-> 전체 폴더에 추가해도 연락처쪽으로 가는 이유는 안에 연락처 폴더가 있기 때문인 것 같음
"""
contacts_folder = namespace.GetDefaultFolder(10)
for folder in contacts_folder.Folders:
    print(folder)

# contacts_folder.Folders.Add("연락처_Test")
# print(contacts_folder)

# contacts_folder.Folders.Remove("연락처_Test")
# 주소록 폴더를 찾습니다.
"""
네임스페이스 뎁스에서 찾으면 안보이게됨 -> 주소록 폴더 하위로 가서 삭제 필요
"""
folder_name = "연락처_Test"
address_book_folder = None
for folder in namespace.Folders:
    if folder.Name == folder_name:
        address_book_folder = folder
        break

# 주소록 폴더를 삭제합니다.
if address_book_folder:
    address_book_folder.Delete()
    print(f"'{folder_name}' 폴더가 삭제되었습니다.")
else:
    print(f"'{folder_name}' 폴더를 찾을 수 없습니다.")

"""
네임스페이스 뎁스에서 찾으면 안보이게됨 -> 주소록 폴더 하위로 가서 삭제 필요
"""
folder_name = "연락처_Test"
address_book_folder = None
for folder in namespace.GetDefaultFolder(10).Folders:
    if folder.Name == folder_name:
        address_book_folder = folder
        break

# 주소록 폴더를 삭제합니다.
if address_book_folder:
    address_book_folder.Delete()
    print(f"'{folder_name}' 폴더가 삭제되었습니다.")
else:
    print(f"'{folder_name}' 폴더를 찾을 수 없습니다.")

"""
삭제후 찾으면? 
하위에 아직 남아있음 -> 폴더 이름이 변경돼서 ex) +(이 컴퓨터만 해당)

전체 삭제하려면 접근해서 다 삭제하기
"""
contacts_folder = namespace.GetDefaultFolder(10)
for folder in contacts_folder.Folders:
    folder.Delete()

"""
주소록 폴더 접근해서 다 삭제
"""
contacts_folder = namespace.GetDefaultFolder(10)
for folder in contacts_folder.Folders:
    folder.Delete()

for folder in contacts_folder.Folders:
    print(folder)

"""

"""