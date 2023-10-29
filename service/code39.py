import os
import csv
from typing import List

import barcode
from barcode.writer import ImageWriter


def generate_code39_images(csv_path: str, dest_dir: str) -> List[str]: 
  print(csv_path, dest_dir)

  with open(csv_path, "r", encoding="utf-8") as f:
    reader = csv.reader(f)
    filenames = []
    for row in reader:
      code_data = row[1].replace('*', '')
      code39 = barcode.get('code39', code_data, writer=ImageWriter(), options={"add_checksum": False})

      os.makedirs(dest_dir, exist_ok=True)
      filename = code39.save(dest_dir + "/" + row[0], options={
        'module_height':25.00,
        'module_width':0.675,
        'text_distance':8.0,
        'font_size':16,
        # 'font_path': os.path.abspath('assets/NotoSans-Regular.ttf')
      })
      filenames.append(filename)
    return filenames
