import openpyxl

# Load the Excel file
workbook = openpyxl.load_workbook('new_test.xlsx')

# Select the desired worksheet
worksheet = workbook['Sheet1']  # Replace 'Sheet1' with the name of your worksheet

# Iterate over each cell in the worksheet
Font_Name_List = []
Font_Size_List = []
for row in worksheet.iter_rows():
    for cell in row:
        # Get the font information for the cell
        font = cell.font
        font_family = font.name
        Font_Name_List.append(font_family)
        font_size = font.size
        Font_Size_List.append(font_size)
        # Print the font family
        # print(f"Font Family: {font.name}")
        # print(f"Font Size: {font.size}")
NoRepeated_Font_Name_List = list(set(Font_Name_List))
NoRepeated_Font_Size_List = list(set(Font_Size_List))
print(Font_Name_List,len(Font_Name_List))
print(Font_Size_List, len(Font_Size_List))
print(NoRepeated_Font_Name_List, len(NoRepeated_Font_Name_List))
print(NoRepeated_Font_Size_List, len(NoRepeated_Font_Size_List))


new_data_name = {}
for k in range(0, len(NoRepeated_Font_Name_List)):
    Font_Name_Number = 0
    for i in range(0, len(Font_Name_List)):        
        if Font_Name_List[i] == f"{NoRepeated_Font_Name_List[k]}":
            Font_Name_Number += 1
    new_data_name[f"{NoRepeated_Font_Name_List[k]}"] = Font_Name_Number
print(new_data_name)

new_data_size = {}
for l in range(0, len(NoRepeated_Font_Size_List)):
    Font_Size_Number = 0
    for i in range(0, len(Font_Name_List)):        
        if Font_Size_List[i] == NoRepeated_Font_Size_List[l]:
            Font_Size_Number += 1
    new_data_size[f"{NoRepeated_Font_Size_List[l]}"] = Font_Size_Number
print(new_data_size)

            
#     if Font_Name_List[i] == f"{NoRepeated_Font_Name_List[2]}":
#         Font_Name_Number += 1
#     if Font_Size_List[i] == 8.0:
#         Font_Size_Number += 1
# print(Font_Name_Number)
# print(Font_Size_Number)
