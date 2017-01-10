#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import comtypes.client
import pptx
import shutil
from PIL import Image

class Ppt_exporter(object):

    def __init__(self, configuration):
        self.presentation_path = configuration.PPT_TEMPLATE
        self.images_export_path = configuration.EXPORT_DIR
        self.file_extension_to_export_slides = configuration.SLIDES_EXPORT_FILE_EXTENSION
        self.params_list = configuration.FULL_PARAMS
        self.date = configuration.TODAY

        self.delete_export_dir(self.images_export_path)
        self.run_presentation_export()
        self.delete_slides(self.presentation_path)

    def run_presentation_export(self):
        self.export_presentation()
        self.adjust_exported_file_names()
        sorted_files = self.sort_files()

        for i, file in enumerate(sorted_files):
            self.decrease_image_size(self.images_export_path + file, 60, self.images_export_path + file[:-4] + '.jpg')
            os.remove(self.images_export_path + file)

        sorted_files = self.sort_files()

        self.split_sorted_files_to_parameter_sub_dirs(sorted_files)

    def decrease_image_size(self, image, decreased_size_percent, export_path):
        im = Image.open(image)
        width, height = im.size

        size_percent = decreased_size_percent

        output_width = width * size_percent / 100
        output_height = height * size_percent / 100

        size = output_width, output_height
        im.thumbnail(size, Image.ANTIALIAS)
        im.save(export_path, "JPEG", dpi=(300, 300))



    def export_presentation(self):
        if not (os.path.isfile(self.presentation_path) and os.path.isdir(self.images_export_path)):
            raise "Please give valid paths for presentation file and slides export path!"

        print "Exporting presentation into .png files"
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = True
        presentation = powerpoint.Presentations.Open(self.presentation_path)
        presentation.Export(self.images_export_path, self.file_extension_to_export_slides)
        presentation.Close()
        powerpoint.Quit()
        print "Export successful"

    def adjust_exported_file_names(self):
        files = os.listdir(self.images_export_path)
        for file in files:
            os.rename(self.images_export_path + file, self.images_export_path + file[6:-4] + '.png')

    def get_sort_key(self, file_name):
        num, extension = file_name.split('.')
        return int(num)

    def sort_files(self):
        files = os.listdir(self.images_export_path)
        files.sort(key=self.get_sort_key)
        return files

    def split_sorted_files_to_parameter_sub_dirs(self, sorted_files):
        sub_lists =  [sorted_files[i:i + 3] for i in xrange(0, len(sorted_files), 3)]
        i = 0
        for list in sub_lists:
            z = 0
            for file in list:
                new_file = str(self.date.year)[2:] + str(self.date.month) + str(self.date.day) +  '_' + self.params_list[i].replace('_', '') + '_' + str(z) + '.jpg'
                os.rename(self.images_export_path + file, self.images_export_path + new_file)
                z += 1
            i += 1

    def delete_export_dir(self, export_dir):
        for the_file in os.listdir(export_dir):
            file_path = os.path.join(export_dir, the_file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception as e:
                "Clearing base dir failed with exception: " + str(e)

    def delete_slides(self, presentation_template):
        presentation = pptx.Presentation(presentation_template)
        xml_slides = presentation.slides._sldIdLst
        slides = list(xml_slides)

        for i in range(0, len(slides)):
            xml_slides.remove(slides[i])

        presentation.save(presentation_template)


from Configuration import Configuration

Ppt_exporter(Configuration())