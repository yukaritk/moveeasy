import os

class CredentialManager:
    def __init__(self, file_path='credentials.txt'):
        self.file_path = file_path

    def save_credentials(self, username, password):
        with open(self.file_path, 'w') as file:
            file.write(f"{username}\n{password}")

    def load_credentials(self):
        if not os.path.exists(self.file_path):
            with open(self.file_path, 'w') as file:
                file.write("")
            return None, None

        with open(self.file_path, 'r') as file:
            lines = file.readlines()
            username = lines[0].strip() if len(lines) > 0 else None
            password = lines[1].strip() if len(lines) > 1 else None
            return username, password
