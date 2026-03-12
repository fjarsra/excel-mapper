# src/models/preset_manager.py
import json
import os

class PresetManager:
    @staticmethod
    def save_preset(filepath, src_file, dest_file, rules):
        """Menyimpan konfigurasi mapping ke format JSON"""
        data = {
            "source_file_template": os.path.basename(src_file),
            "destination_file": os.path.basename(dest_file),
            "rules": rules
        }
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)

    @staticmethod
    def load_preset(filepath):
        """Memuat konfigurasi mapping dari format JSON"""
        if not os.path.exists(filepath):
            raise FileNotFoundError("Preset file tidak ditemukan.")
            
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data