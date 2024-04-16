import os
import shutil
from tqdm import tqdm

class FileList() :
    def __init__(self, file_list) :
        self.file_list = file_list

    def is_in_folder(self, folder_path, reverse=False) :
        
        files_in_folder = os.listdir(folder_path)
        check_dict = {}
        not_in_list = []

        if reverse = False :
            for file in tqdm(self.file_list) :
                if file in files_in_folder :
                    check_dict[file] = True
                else :
                    check_dict[file] = False
            return check_dict

        else :
            for file in tqdm(files_in_folder) :
                if file not in self.file_list :
                    not_in_list.append(file)
                else :
                    continue
            return not_in_list

    def remove_files(self, folder_path, reverse=False, into_bottom=False) :
        
        if reverse=False, into_bottom=False :
            
            files_in_folder = os.listdir(folder_path)

            for file in tqdm(self.file_list) :
                if file in files_in_folder :
                    file_path = os.path.join(folder_path, file)
                    os.remove(file_path)
                else :
                    continue

        elif reverse=False, into_bottom=True :
            for root, dirs, files in tqdm(os.walk(folder_path, topdown=False)) :
                for name in files :
                    if name in self.file_list :
                        file_path = os.path.join(root, name)
                        os.remove(file_path)
                    else :
                        continue

        elif reverse=True, into_bottom=False :

            files_in_folder = os.listdir(folder_path)
            
            for file in tqdm(files_in_folder) :
                if file not in self.file_list :
                    file_path = os.path.join(folder_path, file)
                    os.remove(file_path)
                else :
                    continue

        else :
            for root, dirs, files in tqdm(os.walk(folder_path, topdown=False)) :
                for name in files :
                    if name not in self.file_list :
                        file_path = os.path.join(root, name)
                        os.remove(file_path)
                    else :
                        continue

    def copy_files(self, from_folder, to_folder, reveres=False, into_bottom=False) :
        
        from_folder_files = os.listdir(from_folder)

        if reverse=False, into_bottom=False :
            for file in tqdm(from_folder_files) :

                if file in self.file_list :
                    from_path = os.path.join(from_folder, file)
                    to_path = os.path.join(to_folder, file)
                    shutil.copyfile(from_path, to_path)

                else :
                    continue

        elif reverse=False, into_bottom=True :
            for root, dirs, files in tqdm(os.walk(folder_path, topdown=False)) :
                for name in files :
                    if name in self.file_list :
                        from_path = os.path.join(root, name)
                        to_path = os.path.join(to_folder, name)
                        shutil.copyfile(from_path, to_path)

        elif reverse=True, into_bottom=False :
            for file in tqdm(from_folder_files) :

                if file not in self.file_list :
                    from_path = os.path.join(from_folder, file)
                    to_path = os.path.join(to_folder, file)
                    shutil.copyfile(from_path, to_path)

                else :
                    continue


        else :
            for root, dirs, files in tqdm(os.walk(folder_path, topdown=False)) :
                for name in files :
                    if name not in self.file_list :
                        from_path = os.path.join(root, name)
                        to_path = os.path.join(to_folder, name)
                        shutil.copyfile(from_path, to_path)


    def move_files(self, from_folder, to_folder, reverse=False, into_bottom=False) :
        
        from_folder_files = os.listdir(from_folder)

        if reverse=False, into_bottom=False :
            for file in tqdm(from_folder_files) :

                if file in self.file_list :
                    from_path = os.path.join(from_folder, file)
                    to_path = os.path.join(to_folder, file)
                    shutil.move(from_path, to_path)

                else :
                    continue

        elif reverse=False, into_bottom=True :
            for root, dirs, files in tqdm(os.walk(folder_path, topdown=False)) :
                for name in files :
                    if name in self.file_list :
                        from_path = os.path.join(root, name)
                        to_path = os.path.join(to_folder, name)
                        shutil.move(from_path, to_path)

        elif reverse=True, into_bottom=False :
            for file in tqdm(from_folder_files) :

                if file not in self.file_list :
                    from_path = os.path.join(from_folder, file)
                    to_path = os.path.join(to_folder, file)
                    shutil.move(from_path, to_path)

                else :
                    continue


        else :
            for root, dirs, files in tqdm(os.walk(folder_path, topdown=False)) :
                for name in files :
                    if name not in self.file_list :
                        from_path = os.path.join(root, name)
                        to_path = os.path.join(to_folder, name)
                        shutil.move(from_path, to_path)

    def display_guide(self) :
        print("""
        - is_in_folder(self, folder_path, reverse=False) : 
            reverse=False : file_list안에 있는 파일이 folder_path에 있으면 True, 없으면 False를 value로 하는 딕셔너리를 반환함
            reverse=True : folder_path안에 있지만, file_list안에는 없는 파일들의 리스트를 반환함

        - remove_files(self, folder_path, reverse=False, into_bottom=False) :
            reverse=False : file_list안에 있는 파일이 folder_path에 있으면 이를 삭제함
            reverse=True : folder_path안에 있지만, file_list안에는 없는 파일들을 삭제함
            into_bottom : True = 하위폴더까지 순회하면서 찾음, False : 하위 폴더 순회 안 함

        - copy_files(self, from_folder, to_folder, reveres=False, into_bottom=False) :
            reverse=False : file_list안에 있는 파일이 from_folder에 있으면 이를 to_folder에 복사해서 붙여넣음
            reverse=True : from_folder에는 있지만, file_list안에는 없는 파일들을 to_folder에 복사해서 붙여넣음
            into_bottom : True = 하위폴더까지 순회하면서 찾음, False : 하위 폴더 순회 안 함
        
        - move_files(self, from_folder, to_folder, reverse=False, into_bottom=False) :
            reverse=False : file_list안에 있는 파일이 from_folder에 있으면 이를 to_folder로 이동함
            reverse=True: from_folder에는 있지만, file_list안에는 없는 파일들을 to_folder에 복사해서 붙여 넣음
            into_bottom : True = 하위폴더까지 순회하면서 찾음, False : 하위 폴더 순회 안 함
        """)

    @classmethod
    def execute(cls) :
        lst = []
        instance = cls(lst)
        instance.display_guide()
