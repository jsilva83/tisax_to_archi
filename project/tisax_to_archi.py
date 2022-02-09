# Import external modules.
import pandas as pd

# Import internal modules.
import xls_retrieve_data as xls

# Constants.
TISAX_FILE_NAME = 'C:/Users/jorge.silva/HUFGLOBAL/SGSI ISMS - ISMS Management - ISMS Management/I - TISAX Project/02 '\
                  '- Self assessment/05 - VDA 5_0_4 (05-01-2022)/VDA ISA 5.0.4_EN_archimate.xlsx'
TISAX_SHEET = 'Information Security'
REQUIREMENTS_FILE = 'C:/Users/jorge.silva/Downloads/tisax_requirements.xlsx'


def tisax_to_archi():
    # Create output dataframe.
    # Create an empty data frame.
    df_data = pd.DataFrame(columns=['objective', 'must_requirement', 'sub-requirement'])
    # Get must requirements from TISAX check list.
    a_xls = xls.XlsRetrieveData(in_column=4, in_workbook=TISAX_FILE_NAME, in_worksheet=TISAX_SHEET, in_first_row=5)
    for a_n in range(5, a_xls.last_row + 1):
        # Example:
        # ['1.1.1. The organization needs at least one information security policy. This reflects the importance and
        #          significance of information security and is adapted to the organization. Additional policies may be
        #          appropriate depending on the size and structure of the organization.',
        #   ['+ The requirements for information security have been determined and documented.',
        #    '  - The requirements are adapted to the goals of the organization.',
        #    "  - A policy has been created and approved by the organization's management.",
        #    '+ The policy includes objectives and the significance of information security within the organization.']
        # ]
        a_row = a_xls.get_must_row(a_n)
        if len(a_row) > 0:
            objective = a_row[0]
            requirement = ''
            sub_requirement = ''
            for item in a_row[1]:
                item_list = item.split('+')
                if len(item_list) > 1:
                    requirement = item_list[1].strip()
                item_list = item.split('-')
                if len(item_list) > 1:
                    sub_requirement = item_list[1].strip()
                else:
                    sub_requirement = ''
                df_row_to_add = pd.DataFrame({'objective': [objective], 'must_requirement': [requirement],
                                              'sub-requirement': [sub_requirement]})
                # Add a new row.
                df_data = pd.concat([df_data, df_row_to_add], ignore_index=True)
    # Write Excel file.
    df_data.to_excel(REQUIREMENTS_FILE, index=False)
    return


if __name__ == '__main__':
    tisax_to_archi()
