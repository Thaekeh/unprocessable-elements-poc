import zipfile
import xml.etree.ElementTree as ET

def remove_text_element(pptx_path, slide_number, text_to_remove):
    # Step 1: Unzip the .pptx file
    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        zip_ref.extractall('temp')

    # Step 2: Open the XML file of the target slide
    slide_xml_path = f'temp/ppt/slides/slide{slide_number}.xml'
    tree = ET.parse(slide_xml_path)
    root = tree.getroot()

    # Step 3: Locate and remove the text element
    for text_element in root.iter('t'):
        if text_element.text == text_to_remove:
            text_element.clear()

    # Step 4: Save the modified XML file
    tree.write(slide_xml_path)

    # Step 5: Repackage the .pptx file
    with zipfile.ZipFile('modified.pptx', 'w') as zip_ref:
        for foldername, subfolders, filenames in os.walk('temp'):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, 'temp')
                zip_ref.write(file_path, arcname)

# Example usage
remove_text_element('example.pptx', 1, 'Text to remove')
