# import streamlit as st
# from PIL import Image
# import base64
# from io import BytesIO
# import pandas as pd

# # Function to convert image to base64
# def image_to_base64(img):
#     buffer = BytesIO()
#     img.save(buffer, format="WEBP")
#     img_str = base64.b64encode(buffer.getvalue()).decode()
#     return img_str

# # Set up the app title and subheading
# st.set_page_config(page_title="HASIQ", page_icon=":guardsman:", layout="centered")

# # Load and display the background image
# image = Image.open('qc.webp')
# image_base64 = image_to_base64(image)

# # Initialize session state for page tracking
# if 'page' not in st.session_state:
#     st.session_state.page = 'home'

# # CSS for styling
# st.markdown(
#     f"""
#     <style>
#     .stApp {{
#         background-image: url('data:image/webp;base64,{image_base64}');
#         background-size: cover;
#         background-position: center;
#         height: 100vh;
#         display: flex;
#         justify-content: center;
#         align-items: center;
#         flex-direction: column;
#         margin: 0;
#     }}
#     .content {{
#         text-align: center;
#         color: white;
#         padding: 20px;
#         border-radius: 10px;
#         background: rgba(0, 0, 0, 0.5); /* Optional: Adds a semi-transparent background */
#     }}
#     .title {{
#         font-size: 4em; /* Adjust the size as needed */
#         margin: 0;
#         letter-spacing: 0.5em; /* Reduced spacing between letters */
#     }}
#     .subheader {{
#         font-size: 2em; /* Increased font size */
#         font-style: italic;
#         margin-top: 20px;
#     }}
#     .description {{
#         font-size: 1.2em; /* Increased font size */
#         margin-top: 20px;
#     }}
#     .button {{
#         margin-top: 30px;
#         padding: 10px 20px;
#         font-size: 1.2em;
#         background-color: #007bff;
#         color: white;
#         border: none;
#         border-radius: 5px;
#         cursor: pointer;
#         display: block;
#         margin-left: auto;
#         margin-right: auto;
#     }}
#     .back-button {{
#         margin-top: 20px;
#         padding: 10px 20px;
#         font-size: 1.2em;
#         background-color: #6c757d;
#         color: white;
#         border: none;
#         border-radius: 5px;
#         cursor: pointer;
#     }}
#     </style>
#     """,
#     unsafe_allow_html=True
# )

# # Display content based on the current page
# if st.session_state.page == 'home':
#     st.markdown('''
#         <div class="content">
#             <div class="title">H A S I Q</div>
#             <div class="subheader">Harmony Assured System Interface for Quality</div>
#             <div class="description">Welcome to the HASIQ app. This application helps in ensuring quality through effective system interface and data management. Explore the features and tools provided to achieve seamless quality assurance.</div>
#             <button class="button" id="quality-check-btn">Perform Quality Check</button>
#         </div>
#     ''', unsafe_allow_html=True)
    
#     # Handle button click for page navigation
#     if st.button('Perform Quality Check', key='quality_check'):
#         st.session_state.page = 'quality_check'

# elif st.session_state.page == 'quality_check':
#     # st.markdown('''
#     #        <div class="content">
#     #         <div class="title">Quality Check</div>
            
#     #     </div>
#     # ''', unsafe_allow_html=True)
   
#     st.markdown("""<h1 style="color: white; text-align: center;">QC Check Application</h1>""", unsafe_allow_html=True)    
#     st.markdown("""
#         <style>
#         .stFileUploader label {
#             color: white;font-size: 12px;font-weight: bold;
#         }
#         </style>
#         """, unsafe_allow_html=True)
    
#     uploaded_file1 = st.file_uploader("Upload the Excel file", type=["xlsx"])
    
    
#     # Only proceed if the file is uploaded
#     if uploaded_file1:
        
#         df1 = pd.read_excel(uploaded_file1, sheet_name='SAP',skiprows = 1)  
#         file_name = uploaded_file1.name
#         st.markdown(f"""
#             <div style="background-color: white; color: black; padding: 10px; border-radius: 5px; width: auto; font-size: 14px; border: 1px solid black;">
#                 <strong>Uploaded File:</strong> {file_name}
#             </div>
#             """, unsafe_allow_html=True)
#         df1=df1[['Basic Data 1','Attributes','Values']]
#         df1.fillna('',inplace=True)
#         df1['Values_QC']=None
    
#             # Check if 'Attribute' and 'Value_SP' columns exist in df1
#         if 'Attributes' in df1.columns and 'Values' in df1.columns:
#             # Get Material Number and Material Type from df1
#             material_num_df1 = df1.loc[df1['Attributes'] == "Material", 'Values'].values
#             material_type_df1 = df1.loc[df1['Attributes'] == "Material Type", 'Values'].values
    
#             # Check if Material Number and Material Type exist in df1
#             if len(material_num_df1) > 0 and len(material_type_df1) > 0:
#                 material_num_df1 = material_num_df1[0]
#                 material_type_df1 = material_type_df1[0]
    
#                 # Display the Material Number and Material Type from df1
#                 st.markdown(
#         f'<p class="custom-write">Material Number: {material_num_df1}</p>',
#         unsafe_allow_html=True
#     )
#                 st.markdown(
#                     f'<p class="custom-write">Material Type: {material_type_df1}</p>',
#                     unsafe_allow_html=True
#                 )
                
#                 if df1.loc[df1['Attributes'] == 'Material Type', 'Values'].values == 'ZDP':
    
#                    df2 = pd.read_excel(uploaded_file1, sheet_name='Drug Product Share Point')  # Load the first sheet of file 2
#                    df2=df2[['Attributes','Values']]
#                    # Only proceed if the second sheet exists in the file
#                    if df2 is not None and 'Attributes' in df2.columns and 'Values' in df2.columns:
#                        # Merge dataframes for all attributes
#                        #all_attributes = pd.merge(df1, df2, on='Attribute', how='outer')
#                        df1.loc[df1.loc[df1['Attributes']=='Material'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material']['Values'].values[0]
#                        df1.loc[df1.loc[df1['Attributes']=='Material Type'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Type']['Values'].values[0]
#                        #df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
#                        #Local chemical name
#                        if df2.loc[df2['Attributes']=='Chemical name']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
        
#                        df1.loc[df1.loc[df1['Attributes']=='Material Description'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].values[0]
#                        #df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].str.split().str[0].values[0]
#                        #Crystal Modification
#                        if df2.loc[df2['Attributes']=='Crystal Modification']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].str.split().str[0].values[0]
    
#                        #df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].str.split().str[0].values[0]
#                        #Particle size
#                        if df2.loc[df2['Attributes']=='Particle Size']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].str.split().str[0].values[0]
    
#                        df1.loc[df1['Attributes'] == 'Transport Conditions', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Transport Conditions', 'Values'].str.split().str[0].values[0]
    
#                        #df1.loc[df1['Attributes'] == 'Solvate form', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Solvate Form', 'Values'].str.split().str[0].values[0]
#                        # Solvate 
#                        if df2.loc[df2['Attributes']=='Solvate Form']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='Solvate form'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1['Attributes'] == 'Solvate form', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Solvate Form', 'Values'].str.split().str[0].values[0]
       
#                        #df1.loc[df1['Attributes'] == 'Salt code', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Salt code', 'Values'].str.split().str[0].values[0]
#                        # salt code 
#                        if df2.loc[df2['Attributes']=='Salt code']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='Salt code'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1['Attributes'] == 'Salt code', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Salt code', 'Values'].str.split().str[0].values[0]
       
#                        ##df1.loc[df1.loc[df1['Attributes']=='Categoty'].index,'Values_QC']=df2.loc[df2['Attributes']=='Category']['Values'].str.split().str[0].values[0]
#                        #Category
#                        if not df2.loc[df2['Attributes'] == 'category','Values'].empty:
#                            df1.loc[df1['Attributes'] == 'Category','Values_QC'] = df2.loc[df2['Attributes'] == 'category','Values'].str.split().str[0].values[0]
#                        else:
#                            # Optionally, you can assign a default value if no match is found
#                            df1.loc[df1['Attributes'] == 'Category','Values_QC'] = 'Value not Found in SP'     
#                        df1.loc[df1.loc[df1['Attributes']=='Storage Instructions'].index,'Values_QC']=df2.loc[df2['Attributes']=='Storage conditions']['Values'].values[0]
#                        #df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
#                        #Isotope
    
#                        if df2.loc[df2['Attributes']=='Radio Isotope']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
#                        #df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
#                        #Formula
#                        if df2.loc[df2['Attributes']=='Molecular formula']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
#                        df1.loc[df1.loc[df1['Attributes']=='Release'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Release']['Values'].values[0]
#                        #df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
#                        # CAS Number
#                        if df2.loc[df2['Attributes']=='CAS number']['Values'].isnull().any():
#                            df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']="Value not found in SP"
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
    
#                        ##chewck
#                        df1.loc[df1.loc[df1['Attributes']=='Manufacturing Site'].index,'Values_QC']=df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0]
#                        # z_KSO_Relevant'
#                        listMT = ["ZDS","ZRI"]
#                        if df1.loc[df1['Attributes']=='Material Type','Values_QC'].isin(listMT).any():
#                                df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='X'
#                        else:
#                                df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='Attribute not found in SP'
#                        list1=['ZDS','ZDP','ZPP']
#                        ##aSSUMING THE CONDITION IS SAME IN PRODUCT (AS no. reference)
                    
#                        #AS no. reference
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any():
#                            text=df2.loc[df2['Attributes']=='AS no. reference']['Values'].values[0]
#                            if 'TBD' in text:
#                                df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='TBD'
#                            else:
#                                df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='Value not found in SP'
#                        #Project Code
#                        if (df2['Attributes']=='Project name (Development Code)').any():
#                            df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].str[:6].values[0]
#                        elif (df2['Attributes']=='TRD code').any():
#                            df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].str[:6].values[0]
#                        elif (df2['Attributes']=='Drug TRD abbrevation').any():
#                            df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].str[:6].values[0]
#                        #Manufacturing Line
#                        if not df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].empty:
#                            df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].values[0]
#                        else:
#                            # Optionally, you can assign a default value if no match is found
#                            df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = 'Attribute not found in SP'
#                        #Technology
#                        if not df2.loc[df2['Attributes'] == 'Technology','Values'].empty:
#                            df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = df2.loc[df2['Attributes'] == 'Technology','Values'].values[0]
#                        else:
#                            # Optionally, you can assign a default value if no match is found
#                            df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = 'Value not found in SP'
#                        #Not present in SP
#                        if not df2.loc[df2['Attributes'] == 'Molecular weight','Values'].empty:
#                            df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = df2.loc[df2['Attributes'] == 'Molecular weight','Values'].values[0]
#                        else:
#                            # Optionally, you can assign a default value if no match is found
#                            df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = 'Value not found in SP'
#                        #No mapping
#                        #df1.loc[df1.loc[df1['Attributes']=='Basis Number'].index,'Values_QC']=df2.loc[df2['Attributes']=='Basis Number']['Values'].values[0]
#                        #basis number
#                        if not df2.loc[df2['Attributes'] == 'Basis Number','Values'].empty:
#                            df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = df2.loc[df2['Attributes'] == 'Basis Number','Values'].values[0]
#                        else:
#                            # Optionally, you can assign a default value if no match is found
#                            df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = 'Value not found in SP'
                       
#                        #Total Shelf Life
#                        df1.loc[df1.loc[df1['Attributes']=='Total Shelf life'].index,'Values_QC']=df2.loc[df2['Attributes']=='Shelf life']['Values'].str.split().str[0].values[0]
#                        list2=['ZDS','ZRI','ZEX']
#                        list3=['ZDP','ZPP','ZPM']
#                        #Period indicator for SELF
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
#                            df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='D'
#                        elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
#                                df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='M'
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']=""
                       
#                        #Base Unit of Measure
#                        material_type_value = df1.loc[df1['Attributes'] == 'Material Type','Values_QC'].values[0]
#                        material_description = df1.loc[df1['Attributes'] == 'Material Description','Values_QC'].values[0]
                       
#                        if material_type_value == 'ZDS':
#                            df1.loc[df1['Attributes'] == 'Base Unit of Measure','Values_QC'] = 'KG'
#                        elif material_type_value == 'ZDP':
#                            if 'LIPA' in material_description or 'Solution' in material_description:
#                                df1.loc[df1['Attributes'] == 'Base Unit of Measure','Values_QC'] = 'L'
#                        elif material_type_value == 'ZPP':
#                            if 'LYSI' in material_description:
#                                df1.loc[df1['Attributes'] == 'Base Unit of Measure','Values_QC'] = 'PC'                
                       
                       
         
#                        # #Base Unit of Measure
#                        # if (df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDS').any():
#                        #     df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='KG'
#                        # elif (df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDP').any():
#                        #     if (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIPA').any()) or (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('Solution')).any():
#                        #         df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='L'                
         
#                        #Product Hierarchy
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].isnull().any():
#                             df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].values[0]
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=''
#                        #Co-Marketing
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Co-Marketing']['Values'].isnull().any():
#                             df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=df1.loc[df1['Attributes']=='Co-Marketing']['Values'].values[0]
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=''
#                        #HSE Net Relevant
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].isnull().any():
#                             df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].values[0]
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=''
#                        #Global Assortment Status Code
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].isnull().any():
#                             df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].values[0]
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=''
#                        # if saltcode=' ' && crystal modcifiction = ' ' && then value= project code* etc
#                        # #Conditions not clear assumed
#                        # #015-TRD Synonym Name
#                        # if df1.loc[df1['Attributes']=='Salt code']['Values_QC'].isnull().any():
#                        #     if df1.loc[df1['Attributes']=='Crystal Modification']['Values_QC'].isnull().any():
#                        #         df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=df1.loc[df1['Attributes']=='Project code']['Values'].values[0]
#                        # else:
#                        #     df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=''
                       
                       
#                        #TRD Synonym Nmae
#                        if df1.loc[df1['Attributes'] == 'Salt code', 'Values_QC'].isnull().any():
#                            if df1.loc[df1['Attributes'] == 'Crystal Modification', 'Values_QC'].isnull().any():
#                                # Check if 'Project code' exists and has non-empty values
#                                project_code_values = df1.loc[df1['Attributes'] == 'Project Code', 'Values']
#                                if not project_code_values.empty:
#                                    df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values_QC'] = project_code_values.values[0]
#                                else:
#                                    print("No values found for 'Project code'")
#                        else:
#                            df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values_QC'] = ''
#                        # df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']=df2.loc[df2['Attributes']=='AS no. reference']['Values'].values[0]
#                        # IF MT= ZPP & Descrption has this word LIVI or LISY then value is 21.1 else for all MT 20.2
#                        #Haz Mat number
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                            if df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIVI').any() or df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LISY').any():
#                                df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=21.1
#                        else:
#                                df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=20.2
#                        #Dosage Form
#                        # Update 'Dosage Form' based on 'Material Type'
#                        if df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDP').any():
#                            # Get the value after the second word from 'TRD Synonym Name'
#                            trd_value = df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values'].values
#                            if len(trd_value) > 0:
#                                df1.loc[df1['Attributes'] == 'Dosage Form', 'Values_QC'] = ' '.join(trd_value[0].split()[2:])
#                            else:
#                                df1.loc[df1['Attributes'] == 'Dosage Form', 'Values_QC'] = ''
#                        elif df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDS').any():
#                            # Set 'Dosage Form' to empty
#                            df1.loc[df1['Attributes'] == 'Dosage Form', 'Values_QC'] = ''
#                        else:
#                            # Handle other cases or set default value if needed
#                            pass
#                        #Dosage Strength
#                        # Update 'Dosage Form' based on 'Material Type'
#                        if df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDP').any():
#                            # Get the value after the second word from 'TRD Synonym Name'
#                            trd_value = df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values'].values
#                            if len(trd_value) > 0:
#                                df1.loc[df1['Attributes'] == 'Dosage Strength', 'Values_QC'] = ' '.join(trd_value[0].split()[1:2])
#                            else:
#                                df1.loc[df1['Attributes'] == 'Dosage Strength', 'Values_QC'] = ''
#                        elif df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDS').any():
#                            # Set 'Dosage Form' to empty
#                            df1.loc[df1['Attributes'] == 'Dosage Strength', 'Values_QC'] = ''
#                        else:
#                            # Handle other cases or set default value if needed
#                            pass
#                        #X-Plant matl Status: default value
#                        df1.loc[df1.loc[df1['Attributes']=='X-Plant matl Status'].index,'Values_QC']=df1.loc[df1['Attributes']=='X-Plant matl Status']['Values'].values[0]
#                        #Matl Grp Pack.Matls
#                        df1.loc[df1.loc[df1['Attributes']=='Matl Grp Pack.Matls'].index,'Values_QC']=df1.loc[df1['Attributes']=='Matl Grp Pack.Matls']['Values'].values[0]
#                        #Ref. mat. For pckg
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                            df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']='60000949'
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']=''
#                        #Material Class
#                        df1.loc[df1.loc[df1['Attributes']=='Material class'].index,'Values_QC']=df1.loc[df1['Attributes']=='Material class']['Values'].values[0]
#                        #Batch
#                        df1.loc[df1.loc[df1['Attributes']=='Batch'].index,'Values_QC']=df1.loc[df1['Attributes']=='Batch']['Values'].values[0]
#                        #MRP Type
#                        df1.loc[df1.loc[df1['Attributes']=='MRP Type'].index,'Values_QC']=df1.loc[df1['Attributes']=='MRP Type']['Values'].values[0]
#                        #cosumption mode
#                        df1.loc[df1.loc[df1['Attributes']=='cosumption mode'].index,'Values_QC']=df1.loc[df1['Attributes']=='cosumption mode']['Values'].values[0]
#                        #FWD consumption per.
#                        df1.loc[df1.loc[df1['Attributes']=='FWD consumption per.'].index,'Values_QC']=df1.loc[df1['Attributes']=='FWD consumption per.']['Values'].values[0]
#                        #Bwd consumption per. 
#                        df1.loc[df1.loc[df1['Attributes']=='Bwd consumption per. '].index,'Values_QC']=df1.loc[df1['Attributes']=='Bwd consumption per. ']['Values'].values[0]
#                        #Min. Rem. Shelf life
#                        df1.loc[df1.loc[df1['Attributes']=='Min. Rem. Shelf life '].index,'Values_QC']=df1.loc[df1['Attributes']=='Min. Rem. Shelf life ']['Values'].values[0]
#                        #LE Quantity
#                        df1.loc[df1.loc[df1['Attributes']=='LE Quantity'].index,'Values_QC']=df1.loc[df1['Attributes']=='LE Quantity']['Values'].values[0]
#                        #Un
#                        df1.loc[df1.loc[df1['Attributes']=='Un'].index,'Values_QC']=df1.loc[df1['Attributes']=='Un']['Values'].values[0]
#                        #SuT
#                        df1.loc[df1.loc[df1['Attributes']=='SuT'].index,'Values_QC']=df1.loc[df1['Attributes']=='SuT']['Values'].values[0]
#                        #Inspection setup
#                        df1.loc[df1.loc[df1['Attributes']=='Inspection setup'].index,'Values_QC']=df1.loc[df1['Attributes']=='Inspection setup']['Values'].values[0]
#                        #Rounding rule
#                        df1.loc[df1.loc[df1['Attributes']=='Rounding rule'].index,'Values_QC']=df1.loc[df1['Attributes']=='Rounding rule']['Values'].values[0]
#                        list2=['ZDS','ZRI','ZEX']
#                        list3=['ZDP','ZPP','ZPM']
                       
#                        #Total Shelf Life
#                        df1.loc[df1.loc[df1['Attributes']=='Total Shelf Life'].index,'Values_QC']=df2.loc[df2['Attributes']=='Shelf life']['Values'].str.split().str[0].values[0]
                       
#                        #Period indicator for SELF
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
#                            df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='D'
#                        elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
#                                df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='M'
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']=""
#                        import pandas as pd
#                        import re
                       
#                        # Create a reference dictionary for temperature ranges and their codes
#                        temperature_mapping = {
#                            "15° - 25 °C": "A1",
#                            "15° - 30 °C": "A2",
#                            "20° - 25 °C": "A3",
#                            "Room temperature": "A4",
#                            "≤ 25 °C": "A6",
#                            "≤ 30 °C": "A7",
#                            "<30 °C": "A8",
#                            "2 - 8 °C": "C1",
#                            "< -15 °C": "F1",
#                            "<= -30°C": "F10",
#                            "+36.5°C to 1°C": "F11",
#                            "< 15°C": "F12",
#                            "2 to 28°C": "F13",
#                            "-35°C to -45°C": "F14",
#                            "</= -120°C": "F15",
#                            "<-20°C +/- 5°C": "F16",
#                            "< 8 °C": "F17",
#                            "< 25 °C": "F18",
#                            "< -40 °C": "F2",
#                            "At -20°C ± 5°C": "F3",
#                            "< -60 °C": "F4",
#                            "</= -70 °C": "F5",
#                            "</= -130°C": "F6",
#                            "-10°C to -25°C": "F7",
#                            "-60°C to -90°C": "F8",
#                            "-20 to -10°C": "F9"
#                        }                
         
#                        # Function to clean the input text and map temperature range
#                        def map_storage_conditions(storage_text, mapping):
#                            # Replace unwanted Unicode characters like \u2003 (em space)
#                            cleaned_text = storage_text.replace('\u2003', ' ').strip()
                       
#                            # Try to extract the temperature range using a regex pattern that ignores spaces around °C
#                            match = re.search(r'(\d+\s*(°|º)?\s*(to|[-])\s*\d+\s*(°|º)?\s*(C|°C))', cleaned_text)
                           
#                            if match:
#                                temp_range = match.group(1).strip()
                       
#                                # Normalize spaces, removing spaces around the degree symbol if present
#                                temp_range = re.sub(r'\s*(°|º)\s*C', '°C', temp_range).replace("to", "-").strip()
                       
#                                # Compare against normalized dictionary keys
#                                for key, value in mapping.items():
#                                    normalized_key = re.sub(r'\s*(°|º)\s*C', '°C', key)  # Normalize spaces in dictionary keys
#                                    if temp_range == normalized_key:
#                                        return value
#                                    # Debugging output to show comparison
#                                    # print(f"Checking '{temp_range}' against normalized '{normalized_key}'")
#                            else:
#                                print("No temperature range found.")  # Debugging output
                           
#                            return None
                       
#                        # Extract the storage conditions value from the dataframe
#                        storage_conditions = df2.loc[df2['Attributes'] == 'Storage conditions', 'Values'].values[0]
         
#                        # Map the storage conditions to the corresponding code
#                        mapped_value = map_storage_conditions(storage_conditions, temperature_mapping)
                       
#                        print("Mapped Value:", mapped_value)  # Expected Output: C1
#                        df1.loc[df1.loc[df1['Attributes']=='Stock removal '].index,'Values_QC']=mapped_value
#                        df1.loc[df1.loc[df1['Attributes']=='stock placement'].index,'Values_QC']=mapped_value
#                        #MRP Group
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
#                            if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
#                                df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='ZDP1'
#                            elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                                df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
#                            elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1:
#                                df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
#                            elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1:
#                                df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
#                            elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
#                                df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
#                            else:
#                                df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']=''
#                        #Not present in SP MRP controller
#                        if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
#                            df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='P01'
#                        elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
#                            df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='WKL'
#                        elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                            df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0]
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=''
#                        #Procurement Type
#                        if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='F'
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='E'
#                        #Special Procurement Type
#                        if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='20'
#                        elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='30'
#                        elif df1.loc[df1['Attributes']=='Procurement Type']['Values_QC'].values[0]=='E':
#                            df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
#                        if df1.loc[df1['Attributes']=='Release']['Values'].values[0].upper().find("NOT RELEVANT")!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']='Z0Q001'
#                        elif df1.loc[df1['Attributes']=='Release']['Values'].values[0].upper().find('GLIMS LU')!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']='Z0Q002'
#                        elif df1.loc[df1['Attributes']=='Release']['Values'].values[0].upper().find('GLIMS QA')!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']='Z0Q003'
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']=''
#                        ## valuation class
#                        if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2009'
#                        elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2002'
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']=''
#                        #origin group 
#                        if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
#                            df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q45'
#                        elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 :
#                            df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q35'
#                        else:
#                            df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']=''
#                        ##colour
#                        df1['Values']=df1['Values'].astype(str)
#                        df1['Values_QC']=df1['Values_QC'].astype(str)
#                        df1['Values_cp']=df1['Values'].copy()
#                        df1['Values_QC_cp']=df1['Values_QC'].copy()
#                        df1['Values_cp']=df1['Values_cp'].astype(str)
#                        df1['Values_QC_cp']=df1['Values_QC_cp'].astype(str)
#                        df1['Values_cp']=df1['Values_cp'].apply(str.lower)
#                        df1['Values_QC_cp']=df1['Values_QC_cp'].apply(str.lower)
#                        def remove(text):	
#                            return text.replace(" ","")
#                        df1['Values_cp']=df1['Values_cp'].apply(remove)
#                        df1['Values_QC_cp']=df1['Values_QC_cp'].apply(remove)
#                        df1['Values_cp'].replace("","NA",inplace=True)
#                        df1['Values_QC_cp'].replace("","NA",inplace=True)
#                        df1['Values_cp'].replace('none','NA',inplace=True)
#                        df1['Values_QC_cp'].replace('nan','NA',inplace=True)
#                        df1['Compare']=''
#                        for i in range(0,len(df1)):
#                            df1['Compare']=df1['Values_cp']==df1['Values_QC_cp']
#                        df2=df1.copy()
#                        df2=df2[df2['Compare']==False]
#                        flag=[]
#                        for i in range(0,len(df2)):
#                            if str(df2.iloc[i]['Values_cp']).find(str(df2.iloc[i]['Values_QC_cp']))!=-1 or str(df2.iloc[i]['Values_QC_cp']).find(str(df2.iloc[i]['Values_cp']))!=-1:
#                                 #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"Partial")
#                                 flag.append(2)
#                            else:
#                                 #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"No Match")
#                                 flag.append(0)
#                        df2['Flag']=flag
#                        df3=df1.copy()
#                        df3=df3[df3['Compare']==True]
#                        df3['Flag']=1
#                        final = pd.concat([df3,df2])
#                        temp=pd.merge(df1,final,on=['Attributes','Basic Data 1'],how='inner',suffixes=('','_reight'))
#                        temp=temp[['Basic Data 1', 'Attributes', 'Values', 'Values_QC','Flag']]
#                        h = temp["Values_QC"].replace("nan","Value Not Found in SP",inplace = False)
#                        # def color_selected_columns(row):
#                        #     value = row['Flag']
#                        #     if value ==1:
#                        #         color = 'lightgreen'
#                        #     elif value==2:
#                        #         color = '#ffffe0'
#                        #     else:
#                        #         color = '#ffcccb'
                            
#                        #      # Color only the 'Name' and 'City' columns
#                        #     #return [f'background-color: {color}' if col in ['Values', 'Values_QC'] else '' for col in row.index]
#                        #     return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
#                        #             else 'background-color: white; color: black;' for col in row.index]
#                        def color_selected_columns(row):
#                            value = row['Flag']
                            
#                             # Conditional formatting for 'Values' and 'Values_QC'
#                            if value == 1:
#                                color = 'lightgreen'
#                            elif value == 2:
#                                color = '#ffffe0'
#                            else:
#                                color = '#ffcccb'
                            
#                             # Return style for all columns, including white background for others
#                            return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
#                                    else 'background-color: white; color: black;' for col in row.index]
    
    
#                        #df_styled = temp.style.apply(color_selected_columns, axis=1)
                      
#                        all_attributes = temp
                       
#                        attributes = all_attributes['Attributes'].unique().tolist()
#                        # Create a 'Comparison Status' column
#                        #all_attributes['Comparison Status'] = all_attributes.apply(
#                            #lambda x: 'Pass' if x['Values'] == x['Values_QC'] else 'Fail', axis=1)
         
#                        #selected_id = st.text_input("Enter Attribute to compare", "")
#                        st.markdown("""
#         <div style="background-color: white; color: black; padding: 5px; border-radius: 5px; width: 200px; font-size: 14px; text-align: center;font-weight: bold;">
#             <span>Select Attributes to Compare:</span>
#         </div>
#         """, unsafe_allow_html=True)
#                        selected_id=st.selectbox(
#                             "",
#                             options=attributes,
#                             #default=attributes  # By default, select all attributes
#                         )
#                        # Run the comparison and display the resulting dataframe row for the entered attribute
#                        st.markdown("""
#         <style>
#         /* Custom CSS to make button text bold */
#         .css-1emrehy.edgvbvh3 {
#             font-weight: bold !important; /* Makes button text bold */
#         }
#         </style>
#         """, unsafe_allow_html=True)
#                        if st.button('Run Comparison'):
                           
#                            if selected_id:
#                                if selected_id in all_attributes['Attributes'].values:
#                                    result_row = all_attributes[all_attributes['Attributes'] == selected_id]
#                                    st.markdown(f'<p class="comparison-text">Comparison result for ID: {selected_id}</p>', unsafe_allow_html=True)
#                                    #st.write(f"Comparison result for ID: {selected_id}")
#                                    result=result_row[['Attributes', 'Values', 'Values_QC','Flag']]
#                                    result=pd.DataFrame(result)
#                                    result_styled = result.style.apply(color_selected_columns, axis=1)
#                                    result_styled = result_styled.set_table_styles( [{'selector': 'thead th','props': [('background-color', 'white'), ('color','black'),
#                                    ('border', '1px solid black')]}])
#                                    result_styled = result_styled.hide(axis='index')
                                   
    
#                                    df_styled_html = result_styled.to_html()
                         
#                                    st.markdown(df_styled_html, unsafe_allow_html=True)
                           
#                                    #st.result_row.style.apply(color_selected_columns, axis=1)
#                                else:
                                    
#                                      st.error(f"ID {selected_id} not found in the dataframe.")
#                            else:
#                                st.error("Please enter a valid Attribute.")
         
#                        # Show all attributes button
#                        custom_css = """
#         <style>
#         .comparison-text {
#             color: white;
#         }
#         </style>
#     """         
#                        st.markdown(custom_css, unsafe_allow_html=True)
#                        if st.button('Show all attributes'):
                           
#                            st.markdown('<p class="comparison-text">Comparison result for all attributes</p>', unsafe_allow_html=True)
#                            # Display the comparison result for all attributes
#                            #st.write("Comparison result for all attributes")
#                            at=all_attributes[['Attributes', 'Values', 'Values_QC','Flag']]
#                            at=pd.DataFrame(at)
#                            df_styled = at.style.apply(color_selected_columns, axis=1)
#                            df_styled = df_styled.set_table_styles(
#         [{'selector': 'thead th',
#           'props': [('background-color', 'white'), ('color', 'black'),
#                     ('border', '1px solid black')]}]
                               
#     )
                        
#     # Convert the styled DataFrame to HTML
                          
#                            df_styled = df_styled.hide(axis='index')
#                            df_styled = df_styled.hide(axis='columns', subset=['Flag'])
                           
    
#                            df_styled_html = df_styled.to_html()
                         
#                            st.markdown(df_styled_html, unsafe_allow_html=True)
                           
                           
                           
#                        # Export to Excel button
#                        if st.button("Export to Excel"):
#                            export_path = re.sub(r'\W+', '', str(material_num_df1)) + "_qc_checked.xlsx"
#                            df_styled=all_attributes.style.apply(color_selected_columns, axis=1)
#                            df_styled.to_excel(export_path, index=False)
#                            st.success("Data exported to Excel successfully!")
                       
    
                       
#                    else:
#                        st.error("The second sheet does not exist in the file.")
#                 elif df1.loc[df1['Attributes'] == 'Material Type', 'Values'].values == 'ZDS':
#                       df2 = pd.read_excel(uploaded_file1, sheet_name='Drug Substance Share Point')  # Load the first sheet of file 2
#                       df2=df2[['Attributes','Values']]
#                       # Only proceed if the second sheet exists in the file
#                       if df2 is not None and 'Attributes' in df2.columns and 'Values' in df2.columns:
#                           # Merge dataframes for all attributes
#                           # all_attributes = pd.merge(df1, df2, on='Attribute', how='outer')
#     #########         #########################
#                           df1.loc[df1.loc[df1['Attributes']=='Material'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Material Type'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Type']['Values'].values[0]
#                           #df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
#                           #Local chemical name
#                           if df2.loc[df2['Attributes']=='Chemical name']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
        
#                           df1.loc[df1.loc[df1['Attributes']=='Material Description'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].values[0]
#                           #df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].values[0]
#                           #Crystal Modification
#                           if df2.loc[df2['Attributes']=='Crystal Modification']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].str.split().str[0].values[0]
    
#                           #df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].values[0]
#                           #Particle size
#                           if df2.loc[df2['Attributes']=='Particle Size']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].str.split().str[0].values[0]
    
#                           df1.loc[df1.loc[df1['Attributes']=='Transport Conditions'].index,'Values_QC']=df2.loc[df2['Attributes']=='Transport Conditions']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Storage Instructions'].index,'Values_QC']=df2.loc[df2['Attributes']=='Storage conditions']['Values'].values[0]
#                           #df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
#                           #Isotope
    
#                           if df2.loc[df2['Attributes']=='Radio Isotope']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
#                           #df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
#                           #Formula
#                           if df2.loc[df2['Attributes']=='Molecular formula']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
#                           #df1.loc[df1.loc[df1['Attributes']=='Solvate form'].index,'Values_QC']=df2.loc[df2['Attributes']=='Solvate Form']['Values'].values[0]
#                           # Solvate 
#                           if df2.loc[df2['Attributes']=='Solvate Form']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Solvate form'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1['Attributes'] == 'Solvate form', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Solvate Form', 'Values'].str.split().str[0].values[0]
#                           #df1.loc[df1.loc[df1['Attributes']=='Salt code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Salt code']['Values'].values[0]
#                           # salt code 
#                           if df2.loc[df2['Attributes']=='Salt code']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Salt code'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1['Attributes'] == 'Salt code', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Salt code', 'Values'].str.split().str[0].values[0]
#                           #df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
#                           # CAS Number
#                           if df2.loc[df2['Attributes']=='CAS number']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']="Value not found in SP"
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
    
#                           if (df2['Attributes']=='Project name (Development Code)').any():
#                              df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Project name (Development Code)']['Values'].values[0]
#                           elif (df2['Attributes']=='TRD code').any():
#                               df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='TRD code']['Values'].values[0]
#                           elif (df2['Attributes']=='Drug TRD abbrevation').any():
#                               df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Drug TRD abbrevation']['Values'].values[0]
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=''
                          
#                           #df1.loc[df1.loc[df1['Attributes']=='Molecular weight'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular weight']['Values'].isnull
#                           #Not present in SP
#                           if not df2.loc[df2['Attributes'] == 'Molecular weight','Values'].empty:
#                               df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = df2.loc[df2['Attributes'] == 'Molecular weight','Values'].values[0]
#                           else:
#                               # Optionally, you can assign a default value if no match is found
#                               df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = 'Value not found in SP'
#                           #df1.loc[df1.loc[df1['Attributes']=='Basis Number'].index,'Values_QC']=df2.loc[df2['Attributes']=='Basis Number']['Values'].values[0]
#                           #basis number
#                           if not df2.loc[df2['Attributes'] == 'Basis Number','Values'].empty:
#                               df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = df2.loc[df2['Attributes'] == 'Basis Number','Values'].values[0]
#                           else:
#                               # Optionally, you can assign a default value if no match is found
#                              df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = 'Value not found in SP'
#                           df1.loc[df1.loc[df1['Attributes']=='Total Shelf life'].index,'Values_QC']=df2.loc[df2['Attributes']=='Shelf life']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Total Shelf life'].index,'Values_QC']=df1.loc[df1['Attributes']=='Total Shelf life']['Values_QC'].values[0].replace("days","")
#                           if (df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDS').any():
#                               df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='KG'
#                           elif df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDP':
#                               if (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIVI').any()) or (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('Solution')).any():
#                                   df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='L'
#                           list1=['ZDS','ZDP','ZPP']
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].values[0]
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=''
                              
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Co-Marketing']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=df1.loc[df1['Attributes']=='Co-Marketing']['Values'].values[0]
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=''
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].values[0]
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=''
            
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].isnull().any():
#                               df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].values[0]
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=''
            
#                           if df1.loc[df1['Attributes']=='Salt code']['Values_QC'].isnull().any():
#                               if df1.loc[df1['Attributes']=='Crystal Modification']['Values_QC'].isnull().any():
#                                   df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=df1.loc[df1['Attributes']=='Project code']['Values'].values[0]
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=''
            
#                           ##aSSUMING THE CONDITION IS SAME IN PRODUCT (AS no. reference)
#                           #AS no. reference
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any():
#                               text=df2.loc[df2['Attributes']=='AS no. reference']['Values'].values[0]
#                               if 'TBD' in text:
#                                   df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='TBD'
#                               else:
#                                   df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='Value not found in SP'
            
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                               if df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIVI').any() or df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LISY').any():
#                                   df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=21.1
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=20.2
            
#                           list2=['ZDS','ZRI','ZEX']
       
#                           list3=['ZDP','ZPP','ZPM']
#                           #Period indicator for SELF
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
#                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='D'
#                           elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
#                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='M'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']=""
           
#                           # z_KSO_Relevant'
#                           list4 = ["ZDS","ZRI"]
#                           if df1.loc[df1['Attributes']=='Material Type','Values_QC'].isin(list4).any():
#                                   df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='X'
#                           else:
#                                   df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='Attribute not found in SP'
           
#                           #df1.loc[df1.loc[df1['Attributes']=='Manufacturing Line'].index,'Values_QC']=''
#                           #Manufacturing Line
#                           if not df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].empty:
#                               df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].values[0]
#                           else:
#                               # Optionally, you can assign a default value if no match is found
#                               df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = 'Attribute not found in SP'
#                           #df1.loc[df1.loc[df1['Attributes']=='Technology'].index,'Values_QC']=''
#                           #Technology
#                           if not df2.loc[df2['Attributes'] == 'Technology','Values'].empty:
#                               df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = df2.loc[df2['Attributes'] == 'Technology','Values'].values[0]
#                           else:
#                               # Optionally, you can assign a default value if no match is found
#                               df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = 'Value not found in SP'
#                           df1.loc[df1.loc[df1['Attributes']=='Formualtion Phase'].index,'Values_QC']=''
                          
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
#                               df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=df1.loc[df1['Attributes']=='Dosage Form']['Values']
#                           elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
#                               df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=''  
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=''
       
#                           df1.loc[df1.loc[df1['Attributes']=='X-Plant matl Status'].index,'Values_QC']=df1.loc[df1['Attributes']=='X-Plant matl Status']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Matl Grp Pack.Matls'].index,'Values_QC']=df1.loc[df1['Attributes']=='Matl Grp Pack.Matls']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Material class'].index,'Values_QC']=df1.loc[df1['Attributes']=='Material class']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Batch'].index,'Values_QC']=df1.loc[df1['Attributes']=='Batch']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='MRP Type'].index,'Values_QC']=df1.loc[df1['Attributes']=='MRP Type']['Values'].values[0]
           
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                               df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']='60000949'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=''
           
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                               df1.loc[df1.loc[df1['Attributes']=='Dosage Strength'].index,'Values_QC']='60000949'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Dosage Strength'].index,'Values_QC']=''
                              
           
#                           #Ref. mat. For pckg
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                               df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']='60000949'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']=''
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
#                              if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
#                                 df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='ZDS1'
#                              elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                                  df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
#                              elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1:
#                                  df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
#                              elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1:
#                                  df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
#                              elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
#                                  df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']=''
            
            
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
#                               df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='P01'
#                           elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
#                               df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='WKL'
#                           elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
#                               df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0]
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=''
            
#                           if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                               df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='F'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='E'
#                            #Special Procurement Type
#                           if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
#                               df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='20'
#                           elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                               df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='30'
#                           elif df1.loc[df1['Attributes']=='Procurement Type']['Values_QC'].values[0]=='E':
#                               df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
           
#                           df1.loc[df1.loc[df1['Attributes']=='cosumption mode'].index,'Values_QC']=df1.loc[df1['Attributes']=='cosumption mode']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='FWD consumption per.'].index,'Values_QC']=df1.loc[df1['Attributes']=='FWD consumption per.']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Bwd consumption per. '].index,'Values_QC']=df1.loc[df1['Attributes']=='Bwd consumption per. ']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Min. Rem. Shelf life '].index,'Values_QC']=df1.loc[df1['Attributes']=='Min. Rem. Shelf life ']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Material class'].index,'Values_QC']=df1.loc[df1['Attributes']=='Material class']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Rounding rule'].index,'Values_QC']=df1.loc[df1['Attributes']=='Rounding rule']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='LE Quantity'].index,'Values_QC']='99.999.999,000'
#                           df1.loc[df1.loc[df1['Attributes']=='Un'].index,'Values_QC']=df1.loc[df1['Attributes']=='Un']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='SuT'].index,'Values_QC']=df1.loc[df1['Attributes']=='SuT']['Values'].values[0]
#                           df1.loc[df1.loc[df1['Attributes']=='Inspection setup'].index,'Values_QC']=df1.loc[df1['Attributes']=='Inspection setup']['Values'].values[0]
#                           #QM Authorization
#                           if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
#                               df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q30'
#                           elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 :
#                               df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q25'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']=''
           
#                           if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
#                               df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2009'
#                           elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
#                               df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2002'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']=''
           
#                           #Period indicator for SELF
#                           if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
#                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='D'
#                           elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
#                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='M'
#                           else:
#                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']=""
#                           all_attributes = df1
#                           attributes = all_attributes['Attributes'].unique().tolist()
#                           # Create a 'Comparison Status' column
#                           all_attributes['Comparison Status'] = all_attributes.apply(
#                               lambda x: 'Pass' if x['Values'] == x['Values_QC'] else 'Fail', axis=1)
            
#                           #selected_id = st.text_input("Enter Attribute to compare", "")
#                           df1['Values']=df1['Values'].astype(str)
#                           df1['Values_QC']=df1['Values_QC'].astype(str)
#                           df1['Values_cp']=df1['Values'].copy()
#                           df1['Values_QC_cp']=df1['Values_QC'].copy()
#                           df1['Values_cp']=df1['Values_cp'].astype(str)
#                           df1['Values_QC_cp']=df1['Values_QC_cp'].astype(str)
#                           df1['Values_cp']=df1['Values_cp'].apply(str.lower)
#                           df1['Values_QC_cp']=df1['Values_QC_cp'].apply(str.lower)
#                           def remove(text):	
#                               return text.replace(" ","")
#                           df1['Values_cp']=df1['Values_cp'].apply(remove)
#                           df1['Values_QC_cp']=df1['Values_QC_cp'].apply(remove)
#                           df1['Values_cp'].replace("","NA",inplace=True)
#                           df1['Values_QC_cp'].replace("","NA",inplace=True)
#                           df1['Values_cp'].replace('none','NA',inplace=True)
#                           df1['Values_QC_cp'].replace('nan','NA',inplace=True)
#                           df1['Compare']=''
#                           for i in range(0,len(df1)):
#                               df1['Compare']=df1['Values_cp']==df1['Values_QC_cp']
#                           df2=df1.copy()
#                           df2=df2[df2['Compare']==False]
#                           flag=[]
#                           for i in range(0,len(df2)):
#                               if str(df2.iloc[i]['Values_cp']).find(str(df2.iloc[i]['Values_QC_cp']))!=-1 or str(df2.iloc[i]['Values_QC_cp']).find(str(df2.iloc[i]['Values_cp']))!=-1:
                                  
                                  
                                  
                                  
                                  
#                                     #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"Partial")
#                                   flag.append(2)
#                               else:
#                                     #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"No Match")
#                                   flag.append(0)
#                           df2['Flag']=flag
#                           df3=df1.copy()
#                           df3=df3[df3['Compare']==True]
#                           df3['Flag']=1
#                           final = pd.concat([df3,df2])
#                           temp=pd.merge(df1,final,on=['Attributes','Basic Data 1'],how='inner',suffixes=('','_reight'))
#                           temp=temp[['Basic Data 1', 'Attributes', 'Values', 'Values_QC','Flag']]
#                           h = temp["Values_QC"].replace("nan","Value Not Found in SP",inplace = False)
#                             # def color_selected_columns(row):
#                             #     value = row['Flag']
#                             #     if value ==1:
#                             #         color = 'lightgreen'
#                             #     elif value==2:
#                             #         color = '#ffffe0'
#                             #     else:
#                             #         color = '#ffcccb'
                                
#                             #      # Color only the 'Name' and 'City' columns
#                             #     #return [f'background-color: {color}' if col in ['Values', 'Values_QC'] else '' for col in row.index]
#                             #     return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
#                             #             else 'background-color: white; color: black;' for col in row.index]
#                           def color_selected_columns(row):
                              
                              
#                               value = row['Flag']
                                
#                                 # Conditional formatting for 'Values' and 'Values_QC'
#                               if value == 1:
#                                   color = 'lightgreen'
#                               elif value == 2:
#                                   color = '#ffffe0'
#                               else:
#                                   color = '#ffcccb'
                                
#                                 # Return style for all columns, including white background for others
#                               return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
#                                       else 'background-color: white; color: black;' for col in row.index]
                            
                            
#                             #df_styled = temp.style.apply(color_selected_columns, axis=1)
                            
#                           all_attributes = temp
                            
#                           attributes = all_attributes['Attributes'].unique().tolist()
#                             # Create a 'Comparison Status' column
#                             #all_attributes['Comparison Status'] = all_attributes.apply(
#                                #lambda x: 'Pass' if x['Values'] == x['Values_QC'] else 'Fail', axis=1)
                            
#                             #selected_id = st.text_input("Enter Attribute to compare", "")
#                           st.markdown("""
#                             <div style="background-color: white; color: black; padding: 5px; border-radius: 5px; width: 200px; font-size: 14px; text-align: center;font-weight: bold;">
#                             <span>Select Attributes to Compare:</span>
#                             </div>
#                             """, unsafe_allow_html=True)
#                           selected_id=st.selectbox(
#                                 "",
#                                 options=attributes,
#                                 #default=attributes  # By default, select all attributes
#                             )
#                             # Run the comparison and display the resulting dataframe row for the entered attribute
#                           st.markdown("""
#                             <style>
#                             /* Custom CSS to make button text bold */
#                             .css-1emrehy.edgvbvh3 {
#                             font-weight: bold !important; /* Makes button text bold */
#                             }
#                             </style>
#                             """, unsafe_allow_html=True)
#                           if st.button('Run Comparison'):
                               
#                               if selected_id:
#                                   if selected_id in all_attributes['Attributes'].values:
#                                       result_row = all_attributes[all_attributes['Attributes'] == selected_id]
#                                       st.markdown(f'<p class="comparison-text">Comparison result for ID: {selected_id}</p>', unsafe_allow_html=True)
#                                        #st.write(f"Comparison result for ID: {selected_id}")
#                                       result=result_row[['Attributes', 'Values', 'Values_QC','Flag']]
#                                       result=pd.DataFrame(result)
#                                       result_styled = result.style.apply(color_selected_columns, axis=1)
#                                       result_styled = result_styled.set_table_styles( [{'selector': 'thead th','props': [('background-color', 'white'), ('color','black'),
#                                        ('border', '1px solid black')]}])
#                                       result_styled = result_styled.hide(axis='index')
                                       
                            
#                                       df_styled_html = result_styled.to_html()
                             
#                                       st.markdown(df_styled_html, unsafe_allow_html=True)
                               
#                                        #st.result_row.style.apply(color_selected_columns, axis=1)
#                                   else:
                                        
#                                          st.error(f"ID {selected_id} not found in the dataframe.")
#                               else:
                              
                                    
#                                     st.error("Please enter a valid Attribute.")
                            
#                             # Show all attributes button
#                           custom_css = """
#                             <style>
#                             .comparison-text {
#                             color: white;
#                             }
#                             </style>
#                             """         
#                           st.markdown(custom_css, unsafe_allow_html=True)
#                           if st.button('Show all attributes'):
                              
                              
                               
#                               st.markdown('<p class="comparison-text">Comparison result for all attributes</p>', unsafe_allow_html=True)
#                                # Display the comparison result for all attributes
#                                #st.write("Comparison result for all attributes")
#                               at=all_attributes[['Attributes', 'Values', 'Values_QC','Flag']]
#                               at=pd.DataFrame(at)
#                               df_styled = at.style.apply(color_selected_columns, axis=1)
#                               df_styled = df_styled.set_table_styles(
#                             [{'selector': 'thead th',
#                             'props': [('background-color', 'white'), ('color', 'black'),
#                             ('border', '1px solid black')]}]
                                   
#                             )
                            
#                             # Convert the styled DataFrame to HTML
                              
#                               df_styled = df_styled.hide(axis='index')
#                               df_styled = df_styled.hide(axis='columns', subset=['Flag'])
#                               #df_display = at.drop(columns=['Flag'])
                            
#                               df_styled_html = df_styled.to_html()
                             
#                               st.markdown(df_styled_html, unsafe_allow_html=True)
                               
                               
                               
#                             # Export to Excel button
#                           if st.button("Export to Excel"):
#                               export_path = re.sub(r'\W+', '', str(material_num_df1)) + "_qc_checked.xlsx"
#                               df_styled=all_attributes.style.apply(color_selected_columns, axis=1)
#                               df_styled.to_excel(export_path, index=False)
#                               st.success("Data exported to Excel successfully!")
                            
                            
                            
#                       else:
                              
#                           st.error("The second sheet does not exist in the file.")
                              
#         #                   st.markdown("""
#         # <style>
#         # /* Apply styling to the specific label */
#         # div[data-baseweb="Select"] label {
#         #     color: white;
#         # }
#         # </style>
#         # """, unsafe_allow_html=True)
                          
#         #                   selected_id =st.multiselect(
#         #                     "Select Attributes to Compare",
#         #                     options=['All'] + attributes,
#         #                     #default=attributes  # By default, select all attributes
#         #                  )
#         #                   # Run the comparison and display the resulting dataframe row for the entered attribute
#         #                   if st.button('Run Comparison'):
#         #                       if selected_id:
#         #                           if selected_id in all_attributes['Attributes'].values:
#         #                               result_row = all_attributes[all_attributes['Attributes'] == selected_id]
#         #                               st.write(f"Comparison result for ID: {selected_id}")
#         #                               st.dataframe(result_row[['Attributes', 'Values', 'Values_QC', 'Comparison Status']])
#         #                           else:
#         #                               st.error(f"ID {selected_id} not found in the dataframe.")
#         #                       else:
#         #                           st.error("Please enter a valid Attribute.")
            
#         #                   # Show all attributes button
#         #                   if st.button('Show all attributes'):
#         #                       # Display the comparison result for all attributes
#         #                       st.write("Comparison result for all attributes")
#         #                       st.dataframe(all_attributes[['Attributes', 'Values', 'Values_QC', 'Comparison Status']])
            
#         #                   # Export to Excel button
#         #                   if st.button("Export to Excel"):
#         #                       export_path = re.sub(r'\W+', '', str(material_num_df1)) + "_qc_checked.xlsx"
#         #                       all_attributes.to_excel(export_path, index=False)
#         #                       st.success("Data exported to Excel successfully!")
#         #               else:
#         #                st.error("The second sheet does not exist in the file.")           
#                 else:   
#                     st.error("Material Number and/or Material Type not found in the first sheet.")
#             else:
#                 st.error("The first sheet must contain 'Attribute' and 'Values' columns.")
    
#     # Handle button click for page navigation
#     if st.button('Back', key='back'):
#         st.session_state.page = 'home'


import streamlit as st
from PIL import Image
import base64
from io import BytesIO
import pandas as pd
# Function to convert image to base64
def image_to_base64(img, format="WEBP"):
    buffer = BytesIO()
    img.save(buffer, format=format)
    img_str = base64.b64encode(buffer.getvalue()).decode()
    return img_str

# Set up the app title and subheading
st.set_page_config(page_title="HASIQ", page_icon=":guardsman:", layout="centered")

# Load background images
image_home = Image.open('qc.webp')
image_home_base64 = image_to_base64(image_home, "WEBP")

image_quality_check = Image.open('sw.jpg')
image_quality_check_base64 = image_to_base64(image_quality_check, "JPEG")

# Initialize session state for page tracking
if 'page' not in st.session_state:
    st.session_state.page = 'home'

# CSS for styling
if st.session_state.page == 'home':
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url('data:image/webp;base64,{image_home_base64}');
            background-size: cover;
            background-position: center;
            height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-direction: column;
            margin: 0;
        }}
        .content {{
            text-align: center;
            color: white;
            padding: 20px;
            margin-top: 150px;
            border-radius: 10px;
            background: rgba(0, 0, 0, 0.5); /* Optional: Adds a semi-transparent background */
        }}
        .title {{
            font-size: 4em; /* Adjust the size as needed */
            margin: 0;
            letter-spacing: 0.5em; /* Reduced spacing between letters */
        }}
        .subheader {{
            font-size: 2em; /* Increased font size */
            font-style: italic;
            margin-top: 20px;
        }}
        .description {{
            font-size: 1.2em; /* Increased font size */
            margin-top: 20px;
        }}
        .button-container {{
        text-align: center; /* Centers the button within this container */
        margin-top: 30px; /* Margin for spacing */
}} 
        .button {{
            margin-top: 30px;
            padding: 10px 20px;
            font-size: 1.2em;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            display: block;
            margin-left: auto;
            margin-right: auto;
            text-align: center; 
        }}
        </style>
        
        """,
        unsafe_allow_html=True
    )

    st.markdown('''
        <div class="content">
            <div class="title"style="margin-bottom: 0px;padding-bottom: 0px;line-height: 1;">H A S I Q</div>
            <div class="subheader"style="margin-top: 0px;padding-top: 0px;line-height: 1;">Harmony Assured System Interface for Quality</div>
            <div class="description">Welcome to the HASIQ app. This app is designed to simplify and streamline the quality check process.Get results in under a minute, slashing the time spent on manual checks from 30-40 minutes to just seconds. Explore the features and tools provided to achieve seamless quality assurance.
        </div>
    ''', unsafe_allow_html=True)
    
    # Handle button click for page navigation
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button('Perform Quality Check', key='quality_check'):
            st.session_state.page = 'quality_check'

elif st.session_state.page == 'quality_check':
    
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url('data:image/jpeg;base64,{image_quality_check_base64}');
            background-size: cover;
            background-position: center;
            height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-direction: column;
            margin: 0;
        }}
        .content {{
            text-align: center;
            color: white;
            padding: 20px;
            border-radius: 10px;
            background: rgba(0, 0, 0, 0.5); /* Optional: Adds a semi-transparent background */
        }}
        .title {{
            font-size: 4em; /* Adjust the size as needed */
            margin: 0;
            letter-spacing: 0.5em; /* Reduced spacing between letters */
        }}
        .subheader {{
            font-size: 2em; /* Increased font size */
            font-style: italic;
            margin-top: 20px;
        }}
        .description {{
            font-size: 1.2em; /* Increased font size */
            margin-top: 20px;
        }}
        .button {{
            margin-top: 60px;
            padding: 10px 20px;
            font-size: 1.2em;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            display: block;
            margin-left: auto;
            margin-right: auto;
        }}
        .back-button {{
            margin-top: 20px;
            padding: 10px 20px;
            font-size: 1.2em;
            background-color: #6c757d;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

    # st.markdown('''
    #         <div class="content">
    #          <div class="title">Quality Check</div>
            
    #     </div>
    #  ''', unsafe_allow_html=True)
   
    st.markdown("""<h1 style="color: white; text-align: center;">Quality Check</h1>""", unsafe_allow_html=True)    
    st.markdown("""
        <style>
        .stFileUploader label {
            color: white;font-size: 12px;font-weight: bold;
        }
        </style>
        """, unsafe_allow_html=True)
    st.markdown("""
    <style>
    /* Change the text color of the uploaded file name */
    .uploadedFileName {
        color: white !important;
    }
    /* Change the file uploader font color */
    .stFileUpload {
        color: white !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    uploaded_file1 = st.file_uploader("Upload the Excel file", type=["xlsx"])
    
    import re
    # Only proceed if the file is uploaded
    if uploaded_file1:
        
        df1 = pd.read_excel(uploaded_file1, sheet_name='SAP',skiprows = 1)  
        file_name = uploaded_file1.name
        st.markdown(f"""
            <div style="background-color: white; color: black; padding: 10px; border-radius: 5px; width: auto; font-size: 14px; border: 1px solid black;">
                <strong>Uploaded File:</strong> {file_name}
            </div>
            """, unsafe_allow_html=True)
        df1=df1[['Basic Data 1','Attributes','Values']]
        df1.fillna('',inplace=True)
        df1['Values_QC']=None
    
            # Check if 'Attribute' and 'Value_SP' columns exist in df1
        if 'Attributes' in df1.columns and 'Values' in df1.columns:
            # Get Material Number and Material Type from df1
            material_num_df1 = df1.loc[df1['Attributes'] == "Material", 'Values'].values
            material_type_df1 = df1.loc[df1['Attributes'] == "Material Type", 'Values'].values
    
            # Check if Material Number and Material Type exist in df1
            if len(material_num_df1) > 0 and len(material_type_df1) > 0:
                material_num_df1 = material_num_df1[0]
                material_type_df1 = material_type_df1[0]
    
                # Display the Material Number and Material Type from df1
                st.markdown(
        f'<p class="custom-write"style="color:white;">Material Number: {material_num_df1}</p>',
        unsafe_allow_html=True
    )
                st.markdown(
                    f'<p class="custom-write"style="color:white;">Material Type: {material_type_df1}</p>',
                    unsafe_allow_html=True
                )
                
                if df1.loc[df1['Attributes'] == 'Material Type', 'Values'].values == 'ZDP':
    
                   df2 = pd.read_excel(uploaded_file1, sheet_name='Drug Product Share Point')  # Load the first sheet of file 2
                   df2=df2[['Attributes','Values']]
                   # Only proceed if the second sheet exists in the file
                   if df2 is not None and 'Attributes' in df2.columns and 'Values' in df2.columns:
                       # Merge dataframes for all attributes
                       #all_attributes = pd.merge(df1, df2, on='Attribute', how='outer')
                       df1.loc[df1.loc[df1['Attributes']=='Material'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material']['Values'].values[0]
                       df1.loc[df1.loc[df1['Attributes']=='Material Type'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Type']['Values'].values[0]
                       #df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
                       #Local chemical name
                       if df2.loc[df2['Attributes']=='Chemical name']['Values'].isnull().any():
                           df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
        
                       df1.loc[df1.loc[df1['Attributes']=='Material Description'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].values[0]
                       #df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].str.split().str[0].values[0]
                       #Crystal Modification
                       if df2.loc[df2['Attributes']=='Crystal Modification']['Values'].isnull().any():
                           df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].str.split().str[0].values[0]
    
                       #df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].str.split().str[0].values[0]
                       #Particle size
                       if df2.loc[df2['Attributes']=='Particle Size']['Values'].empty:
                           df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].str.split().str[0].values[0]
    
                       df1.loc[df1['Attributes'] == 'Transport Conditions', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Transport Conditions', 'Values'].str.split().str[0].values[0]
    
                       #df1.loc[df1['Attributes'] == 'Solvate form', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Solvate Form', 'Values'].str.split().str[0].values[0]
                       # Solvate 
                       if df2.loc[df2['Attributes']=='Solvate Form']['Values'].isnull().any():
                           df1.loc[df1.loc[df1['Attributes']=='Solvate form'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1['Attributes'] == 'Solvate form', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Solvate Form', 'Values'].str.split().str[0].values[0]
       
                       #df1.loc[df1['Attributes'] == 'Salt code', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Salt code', 'Values'].str.split().str[0].values[0]
                       # salt code 
                       if df2.loc[df2['Attributes']=='Salt code']['Values'].isnull().any():
                           df1.loc[df1.loc[df1['Attributes']=='Salt code'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1['Attributes'] == 'Salt code', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Salt code', 'Values'].str.split().str[0].values[0]
       
                       ##df1.loc[df1.loc[df1['Attributes']=='Categoty'].index,'Values_QC']=df2.loc[df2['Attributes']=='Category']['Values'].str.split().str[0].values[0]
                       #Category
                       if not df2.loc[df2['Attributes'] == 'category','Values'].isnull().any():
                           df1.loc[df1['Attributes'] == 'Category','Values_QC'] = df2.loc[df2['Attributes'] == 'category','Values'].str.split().str[0].values[0]
                       else:
                           # Optionally, you can assign a default value if no match is found
                           df1.loc[df1['Attributes'] == 'Category','Values_QC'] = 'Value not Found in SP'     
                       df1.loc[df1.loc[df1['Attributes']=='Storage Instructions'].index,'Values_QC']=df2.loc[df2['Attributes']=='Storage conditions']['Values'].values[0]
                       #df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
                       #Isotope
    
                       if df2.loc[df2['Attributes']=='Radio Isotope']['Values'].empty:
                           df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
                       #df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
                       #Formula
                       if df2.loc[df2['Attributes']=='Molecular formula']['Values'].isnull().any():
                           df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
                       df1.loc[df1.loc[df1['Attributes']=='Release'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Release']['Values'].values[0]
                       #df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
                       # CAS Number
                       if df2.loc[df2['Attributes']=='CAS number']['Values'].isnull().any():
                           df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']="Value not found in SP"
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
    
                       ##chewck
                       df1.loc[df1.loc[df1['Attributes']=='Manufacturing Site'].index,'Values_QC']=df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0]
                       # z_KSO_Relevant'
                       listMT = ["ZDS","ZRI"]
                       if df1.loc[df1['Attributes']=='Material Type','Values_QC'].isin(listMT).any():
                               df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='X'
                       else:
                               df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='Attribute not found in SP'
                       list1=['ZDS','ZDP','ZPP']
                       ##aSSUMING THE CONDITION IS SAME IN PRODUCT (AS no. reference)
                    
                       #AS no. reference
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any():
                           text=df2.loc[df2['Attributes']=='AS no. reference']['Values'].values[0]
                           if 'TBD' in text:
                               df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='TBD'
                           else:
                               df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='Value not found in SP'
                       #Project Code
                       if (df2['Attributes']=='Project name (Development Code)').any():
                           df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].str[:6].values[0]
                       elif (df2['Attributes']=='TRD code').any():
                           df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].str[:6].values[0]
                       elif (df2['Attributes']=='Drug TRD abbrevation').any():
                           df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].str[:6].values[0]
                       #Manufacturing Line
                       if not df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].empty:
                           df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].values[0]
                       else:
                           # Optionally, you can assign a default value if no match is found
                           df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = 'Attribute not found in SP'
                       #Technology
                       if not df2.loc[df2['Attributes'] == 'Technology','Values'].isnull().any():
                           df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = df2.loc[df2['Attributes'] == 'Technology','Values'].values[0]
                       else:
                           # Optionally, you can assign a default value if no match is found
                           df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = 'Value not found in SP'
                       #Not present in SP
                       if not df2.loc[df2['Attributes'] == 'Molecular weight','Values'].isnull().any():
                           df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = df2.loc[df2['Attributes'] == 'Molecular weight','Values'].values[0]
                       else:
                           # Optionally, you can assign a default value if no match is found
                           df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = 'Value not found in SP'
                       #No mapping
                       #df1.loc[df1.loc[df1['Attributes']=='Basis Number'].index,'Values_QC']=df2.loc[df2['Attributes']=='Basis Number']['Values'].values[0]
                       #basis number
                       if not df2.loc[df2['Attributes'] == 'Basis Number','Values'].isnull().any():
                           df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = df2.loc[df2['Attributes'] == 'Basis Number','Values'].values[0]
                       else:
                           # Optionally, you can assign a default value if no match is found
                           df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = 'Value not found in SP'
                       
                       #Total Shelf Life
                       df1.loc[df1.loc[df1['Attributes']=='Total Shelf life'].index,'Values_QC']=df2.loc[df2['Attributes']=='Shelf life']['Values'].str.split().str[0].values[0]
                       list2=['ZDS','ZRI','ZEX']
                       list3=['ZDP','ZPP','ZPM']
                       #Period indicator for SELF
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
                           df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='D'
                       elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='M'
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']=""
                       
                       #Base Unit of Measure
                       material_type_value = df1.loc[df1['Attributes'] == 'Material Type','Values_QC'].values[0]
                       material_description = df1.loc[df1['Attributes'] == 'Material Description','Values_QC'].values[0]
                       
                       if material_type_value == 'ZDS':
                           df1.loc[df1['Attributes'] == 'Base Unit of Measure','Values_QC'] = 'KG'
                       elif material_type_value == 'ZDP':
                           if 'LIPA' in material_description or 'Solution' in material_description:
                               df1.loc[df1['Attributes'] == 'Base Unit of Measure','Values_QC'] = 'L'
                       elif material_type_value == 'ZPP':
                           if 'LYSI' in material_description:
                               df1.loc[df1['Attributes'] == 'Base Unit of Measure','Values_QC'] = 'PC'                
                       
                       
         
                       # #Base Unit of Measure
                       # if (df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDS').any():
                       #     df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='KG'
                       # elif (df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDP').any():
                       #     if (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIPA').any()) or (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('Solution')).any():
                       #         df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='L'                
         
                       #Product Hierarchy
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].isnull().any():
                            df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].values[0]
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=''
                       #Co-Marketing
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Co-Marketing']['Values'].isnull().any():
                            df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=df1.loc[df1['Attributes']=='Co-Marketing']['Values'].values[0]
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=''
                       #HSE Net Relevant
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].isnull().any():
                            df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].values[0]
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=''
                       #Global Assortment Status Code
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].isnull().any():
                            df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].values[0]
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=''
                       # if saltcode=' ' && crystal modcifiction = ' ' && then value= project code* etc
                       # #Conditions not clear assumed
                       # #015-TRD Synonym Name
                       # if df1.loc[df1['Attributes']=='Salt code']['Values_QC'].isnull().any():
                       #     if df1.loc[df1['Attributes']=='Crystal Modification']['Values_QC'].isnull().any():
                       #         df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=df1.loc[df1['Attributes']=='Project code']['Values'].values[0]
                       # else:
                       #     df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=''
                       
                       
                       #TRD Synonym Nmae
                       if df1.loc[df1['Attributes'] == 'Salt code', 'Values'].values[0].lower().find("nx")!=-1:
                           if (df1.loc[df1['Attributes'] == 'Crystal Modification', 'Values_QC']=="Value not found in SP").any():
                               # Check if 'Project code' exists and has non-empty values
                               project_code_values = df1.loc[df1['Attributes'] == 'Project Code', 'Values']
                               if not project_code_values.empty:
                                   df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values_QC'] = project_code_values.values[0]
                               else:
                                   print("No values found for 'Project code'")
                       else:
                           df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values_QC'] = ''
                       # df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']=df2.loc[df2['Attributes']=='AS no. reference']['Values'].values[0]
                       # IF MT= ZPP & Descrption has this word LIVI or LISY then value is 21.1 else for all MT 20.2
                       #Haz Mat number
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                           if df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIVI').any() or df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LISY').any():
                               df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=21.1
                       else:
                               df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=20.2
                       #Dosage Form
                       # Update 'Dosage Form' based on 'Material Type'
                       if df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDP').any():
                           # Get the value after the second word from 'TRD Synonym Name'
                           trd_value = df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values'].values
                           if len(trd_value) > 0:
                               df1.loc[df1['Attributes'] == 'Dosage Form', 'Values_QC'] = ' '.join(trd_value[0].split()[2:])
                           else:
                               df1.loc[df1['Attributes'] == 'Dosage Form', 'Values_QC'] = ''
                       elif df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDS').any():
                           # Set 'Dosage Form' to empty
                           df1.loc[df1['Attributes'] == 'Dosage Form', 'Values_QC'] = ''
                       else:
                           # Handle other cases or set default value if needed
                           pass
                       #Dosage Strength
                       # Update 'Dosage Form' based on 'Material Type'
                       if df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDP').any():
                           # Get the value after the second word from 'TRD Synonym Name'
                           trd_value = df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values'].values
                           if len(trd_value) > 0:
                               df1.loc[df1['Attributes'] == 'Dosage Strength', 'Values_QC'] = ' '.join(trd_value[0].split()[1:2])
                           else:
                               df1.loc[df1['Attributes'] == 'Dosage Strength', 'Values_QC'] = ''
                       elif df1.loc[df1['Attributes'] == 'Material Type', 'Values_QC'].eq('ZDS').any():
                           # Set 'Dosage Form' to empty
                           df1.loc[df1['Attributes'] == 'Dosage Strength', 'Values_QC'] = ''
                       else:
                           # Handle other cases or set default value if needed
                           pass
                       #X-Plant matl Status: default value
                       df1.loc[df1.loc[df1['Attributes']=='X-Plant matl Status'].index,'Values_QC']=df1.loc[df1['Attributes']=='X-Plant matl Status']['Values'].values[0]
                       #Matl Grp Pack.Matls
                       df1.loc[df1.loc[df1['Attributes']=='Matl Grp Pack.Matls'].index,'Values_QC']=df1.loc[df1['Attributes']=='Matl Grp Pack.Matls']['Values'].values[0]
                       #Ref. mat. For pckg
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                           df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']='60000949'
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']=''
                       #Material Class
                       df1.loc[df1.loc[df1['Attributes']=='Material class'].index,'Values_QC']=df1.loc[df1['Attributes']=='Material class']['Values'].values[0]
                       #Batch
                       df1.loc[df1.loc[df1['Attributes']=='Batch'].index,'Values_QC']=df1.loc[df1['Attributes']=='Batch']['Values'].values[0]
                       #MRP Type
                       df1.loc[df1.loc[df1['Attributes']=='MRP Type'].index,'Values_QC']=df1.loc[df1['Attributes']=='MRP Type']['Values'].values[0]
                       #cosumption mode
                       df1.loc[df1.loc[df1['Attributes']=='cosumption mode'].index,'Values_QC']=df1.loc[df1['Attributes']=='cosumption mode']['Values'].values[0]
                       #FWD consumption per.
                       df1.loc[df1.loc[df1['Attributes']=='FWD consumption per.'].index,'Values_QC']=df1.loc[df1['Attributes']=='FWD consumption per.']['Values'].values[0]
                       #Bwd consumption per. 
                       df1.loc[df1.loc[df1['Attributes']=='Bwd consumption per. '].index,'Values_QC']=df1.loc[df1['Attributes']=='Bwd consumption per. ']['Values'].values[0]
                       #Min. Rem. Shelf life
                       df1.loc[df1.loc[df1['Attributes']=='Min. Rem. Shelf life '].index,'Values_QC']=df1.loc[df1['Attributes']=='Min. Rem. Shelf life ']['Values'].values[0]
                       #LE Quantity
                       df1.loc[df1.loc[df1['Attributes']=='LE Quantity'].index,'Values_QC']=df1.loc[df1['Attributes']=='LE Quantity']['Values'].values[0]
                       #Un
                       df1.loc[df1.loc[df1['Attributes']=='Un'].index,'Values_QC']=df1.loc[df1['Attributes']=='Un']['Values'].values[0]
                       #SuT
                       df1.loc[df1.loc[df1['Attributes']=='SuT'].index,'Values_QC']=df1.loc[df1['Attributes']=='SuT']['Values'].values[0]
                       #Inspection setup
                       df1.loc[df1.loc[df1['Attributes']=='Inspection setup'].index,'Values_QC']=df1.loc[df1['Attributes']=='Inspection setup']['Values'].values[0]
                       #Rounding rule
                       df1.loc[df1.loc[df1['Attributes']=='Rounding rule'].index,'Values_QC']=df1.loc[df1['Attributes']=='Rounding rule']['Values'].values[0]
                       list2=['ZDS','ZRI','ZEX']
                       list3=['ZDP','ZPP','ZPM']
                       
                       #Total Shelf Life
                       df1.loc[df1.loc[df1['Attributes']=='Total Shelf Life'].index,'Values_QC']=df2.loc[df2['Attributes']=='Shelf life']['Values'].str.split().str[0].values[0]
                       
                       #Period indicator for SELF
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
                           df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='D'
                       elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
                               df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='M'
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']=""
                       import pandas as pd
                       import re
                       
                       # Create a reference dictionary for temperature ranges and their codes
                       temperature_mapping = {
                           "15° - 25 °C": "A1",
                           "15° - 30 °C": "A2",
                           "20° - 25 °C": "A3",
                           "Room temperature": "A4",
                           "≤ 25 °C": "A6",
                           "≤ 30 °C": "A7",
                           "<30 °C": "A8",
                           "2 - 8 °C": "C1",
                           "< -15 °C": "F1",
                           "<= -30°C": "F10",
                           "+36.5°C to 1°C": "F11",
                           "< 15°C": "F12",
                           "2 to 28°C": "F13",
                           "-35°C to -45°C": "F14",
                           "</= -120°C": "F15",
                           "<-20°C +/- 5°C": "F16",
                           "< 8 °C": "F17",
                           "< 25 °C": "F18",
                           "< -40 °C": "F2",
                           "At -20°C ± 5°C": "F3",
                           "< -60 °C": "F4",
                           "</= -70 °C": "F5",
                           "</= -130°C": "F6",
                           "-10°C to -25°C": "F7",
                           "-60°C to -90°C": "F8",
                           "-20 to -10°C": "F9"
                       }                
         
                       # Function to clean the input text and map temperature range
                       def map_storage_conditions(storage_text, mapping):
                           # Replace unwanted Unicode characters like \u2003 (em space)
                           cleaned_text = storage_text.replace('\u2003', ' ').strip()
                       
                           # Try to extract the temperature range using a regex pattern that ignores spaces around °C
                           match = re.search(r'(\d+\s*(°|º)?\s*(to|[-])\s*\d+\s*(°|º)?\s*(C|°C))', cleaned_text)
                           
                           if match:
                               temp_range = match.group(1).strip()
                       
                               # Normalize spaces, removing spaces around the degree symbol if present
                               temp_range = re.sub(r'\s*(°|º)\s*C', '°C', temp_range).replace("to", "-").strip()
                       
                               # Compare against normalized dictionary keys
                               for key, value in mapping.items():
                                   normalized_key = re.sub(r'\s*(°|º)\s*C', '°C', key)  # Normalize spaces in dictionary keys
                                   if temp_range == normalized_key:
                                       return value
                                   # Debugging output to show comparison
                                   # print(f"Checking '{temp_range}' against normalized '{normalized_key}'")
                           else:
                               print("No temperature range found.")  # Debugging output
                           
                           return None
                       
                       # Extract the storage conditions value from the dataframe
                       storage_conditions = df2.loc[df2['Attributes'] == 'Storage conditions', 'Values'].values[0]
         
                       # Map the storage conditions to the corresponding code
                       mapped_value = map_storage_conditions(storage_conditions, temperature_mapping)
                       
                       print("Mapped Value:", mapped_value)  # Expected Output: C1
                       df1.loc[df1.loc[df1['Attributes']=='Stock removal '].index,'Values_QC']=mapped_value
                       df1.loc[df1.loc[df1['Attributes']=='stock placement'].index,'Values_QC']=mapped_value
                       #MRP Group
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
                           if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
                               df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='ZDP1'
                           elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                               df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
                           elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1:
                               df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
                           elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1:
                               df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
                           elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
                               df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
                           else:
                               df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']=''
                       #Not present in SP MRP controller
                       if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
                           df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='P01'
                       elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
                           df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='WKL'
                       elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                           df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0]
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=''
                       #Procurement Type
                       if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='F'
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='E'
                       #Special Procurement Type
                       if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='20'
                       elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='30'
                       elif df1.loc[df1['Attributes']=='Procurement Type']['Values_QC'].values[0]=='E':
                           df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
                       if df1.loc[df1['Attributes']=='Release']['Values'].values[0].upper().find("NOT RELEVANT")!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']='Z0Q001'
                       elif df1.loc[df1['Attributes']=='Release']['Values'].values[0].upper().find('GLIMS LU')!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']='Z0Q002'
                       elif df1.loc[df1['Attributes']=='Release']['Values'].values[0].upper().find('GLIMS QA')!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']='Z0Q003'
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='QM material authorization '].index,'Values_QC']=''
                       ## valuation class
                       if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2009'
                       elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2002'
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']=''
                       #origin group 
                       if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
                           df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q45'
                       elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 :
                           df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q35'
                       else:
                           df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']=''
                       ##colour
                       df1['Values']=df1['Values'].astype(str)
                       df1['Values_QC']=df1['Values_QC'].astype(str)
                       df1['Values_cp']=df1['Values'].copy()
                       df1['Values_QC_cp']=df1['Values_QC'].copy()
                       df1['Values_cp']=df1['Values_cp'].astype(str)
                       df1['Values_QC_cp']=df1['Values_QC_cp'].astype(str)
                       df1['Values_cp']=df1['Values_cp'].apply(str.lower)
                       df1['Values_QC_cp']=df1['Values_QC_cp'].apply(str.lower)
                       def remove(text):	
                           return text.replace(" ","")
                       df1['Values_cp']=df1['Values_cp'].apply(remove)
                       df1['Values_QC_cp']=df1['Values_QC_cp'].apply(remove)
                       df1['Values_cp'].replace("","NA",inplace=True)
                       df1['Values_QC_cp'].replace("","NA",inplace=True)
                       df1['Values_cp'].replace('none','NA',inplace=True)
                       df1['Values_QC_cp'].replace('nan','NA',inplace=True)
                       df1['Compare']=''
                       for i in range(0,len(df1)):
                           df1['Compare']=df1['Values_cp']==df1['Values_QC_cp']
                       df2=df1.copy()
                       df2=df2[df2['Compare']==False]
                       flag=[]
                       for i in range(0,len(df2)):
                           if str(df2.iloc[i]['Values_cp']).find(str(df2.iloc[i]['Values_QC_cp']))!=-1 or str(df2.iloc[i]['Values_QC_cp']).find(str(df2.iloc[i]['Values_cp']))!=-1:
                                #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"Partial")
                                flag.append(2)
                           else:
                                #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"No Match")
                                flag.append(0)
                       df2['Flag']=flag
                       df3=df1.copy()
                       df3=df3[df3['Compare']==True]
                       df3['Flag']=1
                       final = pd.concat([df3,df2])
                       temp=pd.merge(df1,final,on=['Attributes','Basic Data 1'],how='inner',suffixes=('','_reight'))
                       temp=temp[['Basic Data 1', 'Attributes', 'Values', 'Values_QC','Flag']]
                       h = temp["Values_QC"].replace("nan","Value Not Found in SP",inplace = False)
                       # def color_selected_columns(row):
                       #     value = row['Flag']
                       #     if value ==1:
                       #         color = 'lightgreen'
                       #     elif value==2:
                       #         color = '#ffffe0'
                       #     else:
                       #         color = '#ffcccb'
                            
                       #      # Color only the 'Name' and 'City' columns
                       #     #return [f'background-color: {color}' if col in ['Values', 'Values_QC'] else '' for col in row.index]
                       #     return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
                       #             else 'background-color: white; color: black;' for col in row.index]
                       def color_selected_columns(row):
                           value = row['Flag']
                            
                            # Conditional formatting for 'Values' and 'Values_QC'
                           if value == 1:
                               color = 'lightgreen'
                           elif value == 2:
                               color = '#ffbf00'
                           else:
                               color = '#DC143C'
                            
                            # Return style for all columns, including white background for others
                           return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
                                   else 'background-color: white; color: black;' for col in row.index]
    
    
                       #df_styled = temp.style.apply(color_selected_columns, axis=1)
                      
                       all_attributes = temp
                       
                       attributes = all_attributes['Attributes'].unique().tolist()
                       # Create a 'Comparison Status' column
                       #all_attributes['Comparison Status'] = all_attributes.apply(
                           #lambda x: 'Pass' if x['Values'] == x['Values_QC'] else 'Fail', axis=1)
         
                       #selected_id = st.text_input("Enter Attribute to compare", "")
                       st.markdown("""
        <div style="background-color: white; color: black; padding: 5px; border-radius: 5px; width: 200px; font-size: 14px; text-align: center;font-weight: bold;">
            <span>Select Attributes to Compare:</span>
        </div>
        """, unsafe_allow_html=True)
                       selected_id=st.selectbox(
                            "",
                            options=attributes,
                            #default=attributes  # By default, select all attributes
                        )
                       # Run the comparison and display the resulting dataframe row for the entered attribute
                       st.markdown("""
        <style>
        /* Custom CSS to make button text bold */
        .css-1emrehy.edgvbvh3 {
            font-weight: bold !important; /* Makes button text bold */
        }
        </style>
        """, unsafe_allow_html=True)
                       if st.button('Run Comparison'):
                           
                           if selected_id:
                               if selected_id in all_attributes['Attributes'].values:
                                   result_row = all_attributes[all_attributes['Attributes'] == selected_id]
                                   st.markdown(f'<p class="comparison-text"style="font-weight: bold; font-size: 20px;">Comparison result for ID: {selected_id}</p>', unsafe_allow_html=True)
                                   #st.write(f"Comparison result for ID: {selected_id}")
                                   result=result_row[['Attributes', 'Values', 'Values_QC','Flag']]
                                   result=pd.DataFrame(result)
                                   result_styled = result.style.apply(color_selected_columns, axis=1)
                                   result_styled = result_styled.set_table_styles( [{'selector': 'thead th','props': [('background-color', 'white'), ('color','black'),
                                   ('border', '1px solid black')]}])
                                   result_styled = result_styled.hide(axis='index')
                                   
    
                                   df_styled_html = result_styled.to_html()
                         
                                   st.markdown(df_styled_html, unsafe_allow_html=True)
                           
                                   #st.result_row.style.apply(color_selected_columns, axis=1)
                               else:
                                    
                                     st.error(f"ID {selected_id} not found in the dataframe.")
                           else:
                               st.error("Please enter a valid Attribute.")
         
                       # Show all attributes button
                       custom_css = """
        <style>
        .comparison-text {
            color: white;
        }
        </style>
    """         
                       st.markdown(custom_css, unsafe_allow_html=True)
                       if st.button('Show all attributes'):
                           
                           st.markdown('<p class="comparison-text"style="font-weight: bold; font-size: 24px;">Comparison result for all attributes</p>', unsafe_allow_html=True)
                           # Display the comparison result for all attributes
                           #st.write("Comparison result for all attributes")
                           at=all_attributes[['Attributes', 'Values', 'Values_QC','Flag']]
                           at=pd.DataFrame(at)
                           df_styled = at.style.apply(color_selected_columns, axis=1)
                           df_styled = df_styled.set_table_styles(
        [{'selector': 'thead th',
          'props': [('background-color', 'white'), ('color', 'black'),
                    ('border', '1px solid black')]}]
                               
    )
                        
    # Convert the styled DataFrame to HTML
                          
                           df_styled = df_styled.hide(axis='index')
                           df_styled = df_styled.hide(axis='columns', subset=['Flag'])
                           
    
                           df_styled_html = df_styled.to_html()
                         
                           st.markdown(df_styled_html, unsafe_allow_html=True)
                           
                           
                           
                       # Export to Excel button
                       if st.button("Export to Excel"):
                           export_path = re.sub(r'\W+', '', str(material_num_df1)) + "_qc_checked.xlsx"
                           df_styled=all_attributes.style.apply(color_selected_columns, axis=1)
                           df_styled.to_excel(export_path, index=False)
                           st.markdown(
    """
    <style>
    .stAlert p {
        color: white !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)
                           st.success("Data exported to Excel successfully!")
                       
    
                       
                   else:
                       st.error("The second sheet does not exist in the file.")
                elif df1.loc[df1['Attributes'] == 'Material Type', 'Values'].values == 'ZDS':
                      df2 = pd.read_excel(uploaded_file1, sheet_name='Drug Substance Share Point')  # Load the first sheet of file 2
                      df2=df2[['Attributes','Values']]
                      # Only proceed if the second sheet exists in the file
                      if df2 is not None and 'Attributes' in df2.columns and 'Values' in df2.columns:
                          # Merge dataframes for all attributes
                          # all_attributes = pd.merge(df1, df2, on='Attribute', how='outer')
    #########         #########################
                          df1.loc[df1.loc[df1['Attributes']=='Material'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Material Type'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Type']['Values'].values[0]
                          #df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
                          #Local chemical name
                          if df2.loc[df2['Attributes']=='Chemical name']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='031-Local Chemical Name'].index,'Values_QC']=df2.loc[df2['Attributes']=='Chemical name']['Values'].values[0]
        
                          df1.loc[df1.loc[df1['Attributes']=='Material Description'].index,'Values_QC']=df2.loc[df2['Attributes']=='Material Description']['Values'].values[0]
                          #df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].values[0]
                          #Crystal Modification
                          if df2.loc[df2['Attributes']=='Crystal Modification']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Crystal Modification'].index,'Values_QC']=df2.loc[df2['Attributes']=='Crystal Modification']['Values'].str.split().str[0].values[0]
    
                          #df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].values[0]
                          #Particle size
                          if df2.loc[df2['Attributes']=='Particle Size']['Values'].empty:
                              df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Particle Size'].index,'Values_QC']=df2.loc[df2['Attributes']=='Particle size']['Values'].str.split().str[0].values[0]
    
                          df1.loc[df1.loc[df1['Attributes']=='Transport Conditions'].index,'Values_QC']=df2.loc[df2['Attributes']=='Transport Conditions']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Storage Instructions'].index,'Values_QC']=df2.loc[df2['Attributes']=='Storage conditions']['Values'].values[0]
                          #df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
                          #Isotope
    
                          if df2.loc[df2['Attributes']=='Radio Isotope']['Values'].empty:
                              df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Isotope'].index,'Values_QC']=df2.loc[df2['Attributes']=='Radio isotope']['Values'].values[0]
                          #df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
                          #Formula
                          if df2.loc[df2['Attributes']=='Molecular formula']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Formula'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular formula']['Values'].values[0]
                          #df1.loc[df1.loc[df1['Attributes']=='Solvate form'].index,'Values_QC']=df2.loc[df2['Attributes']=='Solvate Form']['Values'].values[0]
                          # Solvate 
                          if df2.loc[df2['Attributes']=='Solvate Form']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='Solvate form'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1['Attributes'] == 'Solvate form', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Solvate Form', 'Values'].str.split().str[0].values[0]
                          #df1.loc[df1.loc[df1['Attributes']=='Salt code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Salt code']['Values'].values[0]
                          # salt code 
                          if df2.loc[df2['Attributes']=='Salt code']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='Salt code'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1['Attributes'] == 'Salt code', 'Values_QC'] = df2.loc[df2['Attributes'] == 'Salt code', 'Values'].str.split().str[0].values[0]
                          #df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
                          # CAS Number
                          if df2.loc[df2['Attributes']=='CAS number']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']="Value not found in SP"
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='CAS number'].index,'Values_QC']=df2.loc[df2['Attributes']=='CAS number']['Values'].values[0]
    
                          if (df2['Attributes']=='Project name (Development Code)').any():
                             df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Project name (Development Code)']['Values'].values[0]
                          elif (df2['Attributes']=='TRD code').any():
                              df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='TRD code']['Values'].values[0]
                          elif (df2['Attributes']=='Drug TRD abbrevation').any():
                              df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=df2.loc[df2['Attributes']=='Drug TRD abbrevation']['Values'].values[0]
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Project Code'].index,'Values_QC']=''
                          
                          #df1.loc[df1.loc[df1['Attributes']=='Molecular weight'].index,'Values_QC']=df2.loc[df2['Attributes']=='Molecular weight']['Values'].isnull
                          #Not present in SP
                          if not df2.loc[df2['Attributes'] == 'Molecular weight','Values'].empty:
                              df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = df2.loc[df2['Attributes'] == 'Molecular weight','Values'].values[0]
                          else:
                              # Optionally, you can assign a default value if no match is found
                              df1.loc[df1['Attributes'] == 'Molecular weight','Values_QC'] = 'Value not found in SP'
                          #df1.loc[df1.loc[df1['Attributes']=='Basis Number'].index,'Values_QC']=df2.loc[df2['Attributes']=='Basis Number']['Values'].values[0]
                          #basis number
                          if not df2.loc[df2['Attributes'] == 'Basis Number','Values'].isnull().any():
                              df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = df2.loc[df2['Attributes'] == 'Basis Number','Values'].values[0]
                          else:
                              # Optionally, you can assign a default value if no match is found
                             df1.loc[df1['Attributes'] == 'Basis Number','Values_QC'] = 'Value not found in SP'
                          df1.loc[df1.loc[df1['Attributes']=='Total Shelf life'].index,'Values_QC']=df2.loc[df2['Attributes']=='Shelf life']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Total Shelf life'].index,'Values_QC']=df1.loc[df1['Attributes']=='Total Shelf life']['Values_QC'].values[0].replace("days","")
                          if (df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDS').any():
                              df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='KG'
                          elif df1.loc[df1['Attributes']=='Material Type']['Values_QC']=='ZDP':
                              if (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIVI').any()) or (df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('Solution')).any():
                                  df1.loc[df1.loc[df1['Attributes']=='Base Unit of Measure'].index,'Values_QC']='L'
                          list1=['ZDS','ZDP','ZPP']
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=df1.loc[df1['Attributes']=='Product Hierarchy']['Values'].values[0]
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Product Hierarchy'].index,'Values_QC']=''
                              
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Co-Marketing']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=df1.loc[df1['Attributes']=='Co-Marketing']['Values'].values[0]
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Co-Marketing'].index,'Values_QC']=''
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=df1.loc[df1['Attributes']=='HSE Net Relevant']['Values'].values[0]
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='HSE Net Relevant'].index,'Values_QC']=''
            
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any() & ~df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].isnull().any():
                              df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=df1.loc[df1['Attributes']=='Global Assortment Status Code']['Values'].values[0]
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Global Assortment Status Code'].index,'Values_QC']=''
            
                          # if df1.loc[df1['Attributes']=='Salt code']['Values_QC'].isnull().any():
                          #     if df1.loc[df1['Attributes']=='Crystal Modification']['Values_QC'].isnull().any():
                          #         df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=df1.loc[df1['Attributes']=='Project code']['Values'].values[0]
                          # else:
                          #     df1.loc[df1.loc[df1['Attributes']=='015-TRD Synonym Name'].index,'Values_QC']=''
                          if df1.loc[df1['Attributes'] == 'Salt code', 'Values'].values[0].lower().find("nx")!=-1:
                              if (df1.loc[df1['Attributes'] == 'Crystal Modification', 'Values_QC']=="Value not found in SP").any():
                           # Check if 'Project code' exists and has non-empty values
                                  project_code_values = df1.loc[df1['Attributes'] == 'Project Code', 'Values']
                                  if not project_code_values.empty:
                                      df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values_QC'] = project_code_values.values[0]
                                  else:
                                      print("No values found for 'Project code'")
                          else:
                              df1.loc[df1['Attributes'] == '015-TRD Synonym Name', 'Values_QC'] = ''
            
                          ##aSSUMING THE CONDITION IS SAME IN PRODUCT (AS no. reference)
                          #AS no. reference
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list1).any():
                              text=df2.loc[df2['Attributes']=='AS no. reference']['Values'].values[0]
                              if 'TBD' in text:
                                  df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='TBD'
                              else:
                                  df1.loc[df1.loc[df1['Attributes']=='Analytical Specification No'].index,'Values_QC']='Value not found in SP'
            
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                              if df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LIVI').any() or df1.loc[df1['Attributes']=='Material Description']['Values_QC'].str.contains('LISY').any():
                                  df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=21.1
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Haz. Material number'].index,'Values_QC']=20.2
            
                          list2=['ZDS','ZRI','ZEX']
       
                          list3=['ZDP','ZPP','ZPM']
                          #Period indicator for SELF
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
                              df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='D'
                          elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
                              df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']='M'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELF'].index,'Values_QC']=""
           
                          # z_KSO_Relevant'
                          list4 = ["ZDS","ZRI"]
                          if df1.loc[df1['Attributes']=='Material Type','Values_QC'].isin(list4).any():
                                  df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='X'
                          else:
                                  df1.loc[df1.loc[df1['Attributes']=='z_KSO_Relevant'].index,'Values_QC']='Attribute not found in SP'
           
                          #df1.loc[df1.loc[df1['Attributes']=='Manufacturing Line'].index,'Values_QC']=''
                          #Manufacturing Line
                          if not df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].empty:
                              df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = df2.loc[df2['Attributes'] == 'Manufacturing Line','Values'].values[0]
                          else:
                              # Optionally, you can assign a default value if no match is found
                              df1.loc[df1['Attributes'] == 'Manufacturing Line','Values_QC'] = 'Attribute not found in SP'
                          #df1.loc[df1.loc[df1['Attributes']=='Technology'].index,'Values_QC']=''
                          #Technology
                          if not df2.loc[df2['Attributes'] == 'Technology','Values'].isnull().any():
                              df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = df2.loc[df2['Attributes'] == 'Technology','Values'].values[0]
                          else:
                              # Optionally, you can assign a default value if no match is found
                              df1.loc[df1['Attributes'] == 'Technology','Values_QC'] = 'Value not found in SP'
                          df1.loc[df1.loc[df1['Attributes']=='Formualtion Phase'].index,'Values_QC']=''
                          
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
                              df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=df1.loc[df1['Attributes']=='Dosage Form']['Values']
                          elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
                              df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=''  
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=''
       
                          df1.loc[df1.loc[df1['Attributes']=='X-Plant matl Status'].index,'Values_QC']=df1.loc[df1['Attributes']=='X-Plant matl Status']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Matl Grp Pack.Matls'].index,'Values_QC']=df1.loc[df1['Attributes']=='Matl Grp Pack.Matls']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Material class'].index,'Values_QC']=df1.loc[df1['Attributes']=='Material class']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Batch'].index,'Values_QC']=df1.loc[df1['Attributes']=='Batch']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='MRP Type'].index,'Values_QC']=df1.loc[df1['Attributes']=='MRP Type']['Values'].values[0]
           
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                              df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']='60000949'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Dosage Form'].index,'Values_QC']=''
           
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                              df1.loc[df1.loc[df1['Attributes']=='Dosage Strength'].index,'Values_QC']='60000949'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Dosage Strength'].index,'Values_QC']=''
                          #Category
                          if not df2.loc[df2['Attributes'] == 'category','Values'].isnull().any():
                              df1.loc[df1['Attributes'] == 'Category','Values_QC'] = df2.loc[df2['Attributes'] == 'category','Values'].str.split().str[0].values[0]
                          else:
                              # Optionally, you can assign a default value if no match is found
                              df1.loc[df1['Attributes'] == 'Category','Values_QC'] = 'Value not Found in SP'
                              
           
                          #Ref. mat. For pckg
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                              df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']='60000949'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Ref. mat. For pckg'].index,'Values_QC']=''
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
                             if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
                                df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='ZDS1'
                             elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                                 df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
                             elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1:
                                 df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXT'
                             elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1:
                                 df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
                             elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
                                 df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']='CEXE'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='MRP Group'].index,'Values_QC']=''
            
            
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDS':
                              df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='P01'
                          elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZDP':
                              df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']='WKL'
                          elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].values[0]=='ZPP':
                              df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0]
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='MRP Controller'].index,'Values_QC']=''
            
                          if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                              df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='F'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Procurement Type'].index,'Values_QC']='E'
                           #Special Procurement Type
                          if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("others")!=-1:
                              df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='20'
                          elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                              df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']='30'
                          elif df1.loc[df1['Attributes']=='Procurement Type']['Values_QC'].values[0]=='E':
                              df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Special Procurement Type'].index,'Values_QC']=''
           
                          df1.loc[df1.loc[df1['Attributes']=='cosumption mode'].index,'Values_QC']=df1.loc[df1['Attributes']=='cosumption mode']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='FWD consumption per.'].index,'Values_QC']=df1.loc[df1['Attributes']=='FWD consumption per.']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Bwd consumption per. '].index,'Values_QC']=df1.loc[df1['Attributes']=='Bwd consumption per. ']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Min. Rem. Shelf life '].index,'Values_QC']=df1.loc[df1['Attributes']=='Min. Rem. Shelf life ']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Material class'].index,'Values_QC']=df1.loc[df1['Attributes']=='Material class']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Rounding rule'].index,'Values_QC']=df1.loc[df1['Attributes']=='Rounding rule']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='LE Quantity'].index,'Values_QC']=df1.loc[df1['Attributes']=='LE Quantity']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Un'].index,'Values_QC']=df1.loc[df1['Attributes']=='Un']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='SuT'].index,'Values_QC']=df1.loc[df1['Attributes']=='SuT']['Values'].values[0]
                          df1.loc[df1.loc[df1['Attributes']=='Inspection setup'].index,'Values_QC']=df1.loc[df1['Attributes']=='Inspection setup']['Values'].values[0]
                          #QM Authorization
                          if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
                              df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q30'
                          elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("external")!=-1 or df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("techops")!=-1 :
                              df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']='0Q25'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Origin group'].index,'Values_QC']=''
           
                          if df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("own")!=-1:
                              df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2009'
                          elif df2.loc[df2['Attributes']=='Manufacturing']['Values'].values[0].lower().find("nto")!=-1:
                              df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']='2002'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Valuation class'].index,'Values_QC']=''
           
                          #Period indicator for SELF
                          if df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list2).any():
                              df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='D'
                          elif df1.loc[df1['Attributes']=='Material Type']['Values_QC'].isin(list3).any():
                              df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']='M'
                          else:
                              df1.loc[df1.loc[df1['Attributes']=='Period indicator for SELD'].index,'Values_QC']=""
                          all_attributes = df1
                          attributes = all_attributes['Attributes'].unique().tolist()
                          # Create a 'Comparison Status' column
                          all_attributes['Comparison Status'] = all_attributes.apply(
                              lambda x: 'Pass' if x['Values'] == x['Values_QC'] else 'Fail', axis=1)
                          def extract_substrings(text, chunk_size=2):
    # Use regular expression to find any words or numbers as chunks
                              words = re.findall(r'\w+', text)
                              substrings = set()
                              for word in words:
                                  for i in range(len(word) - chunk_size + 1):
                                      substrings.add(word[i:i + chunk_size])
                              return substrings
 
# Modified partial matching function
                          def partial_match(row):
                              values_cp_substrings = extract_substrings(row['Values_cp'])
                              values_qc_cp_substrings = extract_substrings(row['Values_QC_cp'])
                           
                              # Check if there's any intersection between substrings
                              if values_cp_substrings.intersection(values_qc_cp_substrings):
                                  return 2  # Partial match
                              else:
                                  return 0  # No match
                                      
                          #selected_id = st.text_input("Enter Attribute to compare", "")
                          df1['Values']=df1['Values'].astype(str)
                          df1['Values_QC']=df1['Values_QC'].astype(str)
                          df1['Values_cp']=df1['Values'].copy()
                          df1['Values_QC_cp']=df1['Values_QC'].copy()
                          df1['Values_cp']=df1['Values_cp'].astype(str)
                          df1['Values_QC_cp']=df1['Values_QC_cp'].astype(str)
                          df1['Values_cp']=df1['Values_cp'].apply(str.lower)
                          df1['Values_QC_cp']=df1['Values_QC_cp'].apply(str.lower)
                          def remove(text):	
                              return text.replace(" ","")
                          df1['Values_cp']=df1['Values_cp'].apply(remove)
                          df1['Values_QC_cp']=df1['Values_QC_cp'].apply(remove)
                          df1['Values_cp'].replace("","NA",inplace=True)
                          df1['Values_QC_cp'].replace("","NA",inplace=True)
                          df1['Values_cp'].replace('none','NA',inplace=True)
                          df1['Values_QC_cp'].replace('nan','NA',inplace=True)
                          df1['Compare']=''
                          for i in range(0,len(df1)):
                              df1['Compare']=df1['Values_cp']==df1['Values_QC_cp']
                          df2=df1.copy()
                          df2=df2[df2['Compare']==False]
                          flag=[]
                          for i in range(0,len(df2)):
                              if str(df2.iloc[i]['Values_cp']).find(str(df2.iloc[i]['Values_QC_cp']))!=-1 or str(df2.iloc[i]['Values_QC_cp']).find(str(df2.iloc[i]['Values_cp']))!=-1:
                                  
                                  
                                  
                                  
                                  
                                    #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"Partial")
                                  flag.append(2)
                              else:
                                    #print(df2.iloc[i]['Values_cp'],df2.iloc[i]['Values_QC_cp'],"No Match")
                                  flag.append(0)   #partial_match(df2.iloc[i])
                          df2['Flag']=flag
                          df3=df1.copy()
                          df3=df3[df3['Compare']==True]
                          df3['Flag']=1
                          final = pd.concat([df3,df2])
                          temp=pd.merge(df1,final,on=['Attributes','Basic Data 1'],how='inner',suffixes=('','_reight'))
                          temp=temp[['Basic Data 1', 'Attributes', 'Values', 'Values_QC','Flag']]
                          h = temp["Values_QC"].replace("nan","Value Not Found in SP",inplace = False)
                            # def color_selected_columns(row):
                            #     value = row['Flag']
                            #     if value ==1:
                            #         color = 'lightgreen'
                            #     elif value==2:
                            #         color = '#ffffe0'
                            #     else:
                            #         color = '#ffcccb'
                                
                            #      # Color only the 'Name' and 'City' columns
                            #     #return [f'background-color: {color}' if col in ['Values', 'Values_QC'] else '' for col in row.index]
                            #     return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
                            #             else 'background-color: white; color: black;' for col in row.index]
                          def color_selected_columns(row):
                              
                              
                              value = row['Flag']
                                
                                # Conditional formatting for 'Values' and 'Values_QC'
                              if value == 1:
                                  color = 'lightgreen'
                              elif value == 2:
                                  color = '#ffbf00'
                              else:
                                  color = '#DC143C'
                                
                                # Return style for all columns, including white background for others
                              return [f'background-color: {color}; color: black;' if col in ['Values', 'Values_QC'] 
                                      else 'background-color: white; color: black;' for col in row.index]
                            
                            
                            #df_styled = temp.style.apply(color_selected_columns, axis=1)
                            
                          all_attributes = temp
                            
                          attributes = all_attributes['Attributes'].unique().tolist()
                            # Create a 'Comparison Status' column
                            #all_attributes['Comparison Status'] = all_attributes.apply(
                               #lambda x: 'Pass' if x['Values'] == x['Values_QC'] else 'Fail', axis=1)
                            
                            #selected_id = st.text_input("Enter Attribute to compare", "")
                          st.markdown("""
                            <div style="background-color: white; color: black; padding: 5px; border-radius: 5px; width: 200px; font-size: 14px; text-align: center;font-weight: bold;">
                            <span>Select Attributes to Compare:</span>
                            </div>
                            """, unsafe_allow_html=True)
                          selected_id=st.selectbox(
                                "",
                                options=attributes,
                                #default=attributes  # By default, select all attributes
                            )
                            # Run the comparison and display the resulting dataframe row for the entered attribute
                          st.markdown("""
                            <style>
                            /* Custom CSS to make button text bold */
                            .css-1emrehy.edgvbvh3 {
                            font-weight: bold !important; /* Makes button text bold */
                            }
                            </style>
                            """, unsafe_allow_html=True)
                          if st.button('Run Comparison'):
                               
                              if selected_id:
                                  if selected_id in all_attributes['Attributes'].values:
                                      result_row = all_attributes[all_attributes['Attributes'] == selected_id]
                                      st.markdown(f'<p class="comparison-text"style="font-weight: bold; font-size: 20px;">Comparison result for ID: {selected_id}</p>', unsafe_allow_html=True)
                                       #st.write(f"Comparison result for ID: {selected_id}")
                                      result=result_row[['Attributes', 'Values', 'Values_QC','Flag']]
                                      result=pd.DataFrame(result)
                                      result_styled = result.style.apply(color_selected_columns, axis=1)
                                      result_styled = result_styled.set_table_styles( [{'selector': 'thead th','props': [('background-color', 'white'), ('color','black'),
                                       ('border', '1px solid black')]}])
                                      result_styled = result_styled.hide(axis='index')
                                       
                            
                                      df_styled_html = result_styled.to_html()
                             
                                      st.markdown(df_styled_html, unsafe_allow_html=True)
                               
                                       #st.result_row.style.apply(color_selected_columns, axis=1)
                                  else:
                                        
                                         st.error(f"ID {selected_id} not found in the dataframe.")
                              else:
                              
                                    
                                    st.error("Please enter a valid Attribute.")
                            
                            # Show all attributes button
                          custom_css = """
                            <style>
                            .comparison-text {
                            color: white;
                            }
                            </style>
                            """         
                          st.markdown(custom_css, unsafe_allow_html=True)
                          if st.button('Show all attributes'):
                              
                              
                               
                              st.markdown('<p class="comparison-text"style="font-weight: bold; font-size: 24px;">Comparison result for all attributes</p>', unsafe_allow_html=True)
                               # Display the comparison result for all attributes
                               #st.write("Comparison result for all attributes")
                              at=all_attributes[['Attributes', 'Values', 'Values_QC','Flag']]
                              at=pd.DataFrame(at)
                              df_styled = at.style.apply(color_selected_columns, axis=1)
                              df_styled = df_styled.set_table_styles(
                            [{'selector': 'thead th',
                            'props': [('background-color', 'white'), ('color', 'black'),
                            ('border', '1px solid black')]}]
                                   
                            )
                            
                            # Convert the styled DataFrame to HTML
                              
                              df_styled = df_styled.hide(axis='index')
                              df_styled = df_styled.hide(axis='columns', subset=['Flag'])
                              #df_display = at.drop(columns=['Flag'])
                            
                              df_styled_html = df_styled.to_html()
                             
                              st.markdown(df_styled_html, unsafe_allow_html=True)
                               
                               
                               
                            # Export to Excel button
                          if st.button("Export to Excel"):
                              export_path = re.sub(r'\W+', '', str(material_num_df1)) + "_qc_checked.xlsx"
                              df_styled=all_attributes.style.apply(color_selected_columns, axis=1)
                              df_styled.to_excel(export_path, index=False)
                              st.markdown(
    """
    <style>
    .stAlert p {
        color: white !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)
                              st.success("Data exported to Excel successfully!")
                            
                            
                            
                      else:
                              
                          st.error("The second sheet does not exist in the file.")
                              
        #                   st.markdown("""
        # <style>
        # /* Apply styling to the specific label */
        # div[data-baseweb="Select"] label {
        #     color: white;
        # }
        # </style>
        # """, unsafe_allow_html=True)
                          
        #                   selected_id =st.multiselect(
        #                     "Select Attributes to Compare",
        #                     options=['All'] + attributes,
        #                     #default=attributes  # By default, select all attributes
        #                  )
        #                   # Run the comparison and display the resulting dataframe row for the entered attribute
        #                   if st.button('Run Comparison'):
        #                       if selected_id:
        #                           if selected_id in all_attributes['Attributes'].values:
        #                               result_row = all_attributes[all_attributes['Attributes'] == selected_id]
        #                               st.write(f"Comparison result for ID: {selected_id}")
        #                               st.dataframe(result_row[['Attributes', 'Values', 'Values_QC', 'Comparison Status']])
        #                           else:
        #                               st.error(f"ID {selected_id} not found in the dataframe.")
        #                       else:
        #                           st.error("Please enter a valid Attribute.")
            
        #                   # Show all attributes button
        #                   if st.button('Show all attributes'):
        #                       # Display the comparison result for all attributes
        #                       st.write("Comparison result for all attributes")
        #                       st.dataframe(all_attributes[['Attributes', 'Values', 'Values_QC', 'Comparison Status']])
            
        #                   # Export to Excel button
        #                   if st.button("Export to Excel"):
        #                       export_path = re.sub(r'\W+', '', str(material_num_df1)) + "_qc_checked.xlsx"
        #                       all_attributes.to_excel(export_path, index=False)
        #                       st.success("Data exported to Excel successfully!")
        #               else:
        #                st.error("The second sheet does not exist in the file.")           
                else:   
                    st.error("Material Number and/or Material Type not found in the first sheet.")
            else:
                st.error("The first sheet must contain 'Attribute' and 'Values' columns.")
    
    # Handle button click for page navigation
    if st.button('Back', key='back'):
        st.session_state.page = 'home'