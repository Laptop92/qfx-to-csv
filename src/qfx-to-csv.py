import os
import re
import xml.etree.ElementTree as ET

qfx_path = "."
csv_path = "."

qfx_file_extension = ".qfx"
csv_file_extension = ".csv"

qfx_regex = re.compile(re.escape(qfx_file_extension), re.IGNORECASE)

def extract_XML_string(file_path):
    with open(file_path) as f:
        xml_string = f.read()
    
    index_of_first_angle_bracket = xml_string.find("<")
    
    if index_of_first_angle_bracket > 0:
        xml_string = xml_string[index_of_first_angle_bracket:]
    
    return xml_string

def get_child_by_tag(node, target_tag):
    for child in node:
        if child.tag == target_tag:
            return child
        
    raise AttributeError("Unable to find a child of " + node.tag + 
                             " named " + target_tag)
    
def get_column_headers(INVTRANLIST_node):
    first_transaction_node = get_child_by_tag(INVTRANLIST_node, "INVBANKTRAN")
    
    STMTTRN_node = get_child_by_tag(first_transaction_node, "STMTTRN")
    
    column_headers = ["SUBACCTFUND"]
    
    for child in STMTTRN_node:
        column_headers.append(child.tag)
    
    return column_headers

def convert_to_row(transaction_node, column_headers):
    SUBACCTFUND_node = get_child_by_tag(transaction_node, "SUBACCTFUND")
    
    row_dict = {
        "SUBACCTFUND": SUBACCTFUND_node.text
    }
    
    STMTTRN_node = get_child_by_tag(transaction_node, "STMTTRN")
    
    for child in STMTTRN_node:
        row_dict[child.tag] = child.text
        
    def get_cell_value(column_header):
        cell_value = row_dict[column_header]
        
        if cell_value is None:
            cell_value = ""
            
        return cell_value
    
    cell_values = map(get_cell_value, column_headers) 
    row = ",".join(cell_values) + "\n"
    
    return row

def convert_to_csv(qfx_file_path):
    print("Converting " + qfx_file_path)

    XML_string = extract_XML_string(qfx_file_path)
    
    root = ET.fromstring(XML_string)
    
    INVSTMTMSGSRSV1_node = get_child_by_tag(root, "INVSTMTMSGSRSV1")
    INVSTMTTRNRS_node = get_child_by_tag(INVSTMTMSGSRSV1_node, "INVSTMTTRNRS")
    INVSTMTRS_node = get_child_by_tag(INVSTMTTRNRS_node, "INVSTMTRS")
    INVTRANLIST_node = get_child_by_tag(INVSTMTRS_node, "INVTRANLIST")
    
    qfx_file_name = os.path.basename(qfx_file_path)
    csv_file_name = qfx_regex.sub(csv_file_extension, qfx_file_name)
    
    csv_file_path = os.path.join(csv_path, csv_file_name)
    
    with open(csv_file_path, 'w') as csv_file:
        column_headers = get_column_headers(INVTRANLIST_node)
        header = ",".join(column_headers) + "\n"
        
        csv_file.write(header)
        
        transaction_nodes = filter(lambda node: node.tag == "INVBANKTRAN", 
                                       INVTRANLIST_node)
        
        for transaction_node in transaction_nodes:
            row = convert_to_row(transaction_node, column_headers)
            
            csv_file.write(row)
    
    print("Saved output to " + csv_file_path)

def main():
    file_names_in_qfx_directory = os.listdir(qfx_path)
    
    for file_name in file_names_in_qfx_directory:
        if file_name.lower().endswith(qfx_file_extension):
            qfx_file_path = os.path.join(qfx_path, file_name)
            
            try:
                convert_to_csv(qfx_file_path)
            except Exception as e:
                print(e)
    
if __name__ == "__main__":
    main()
