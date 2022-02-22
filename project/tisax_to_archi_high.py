# Import external modules.
import pandas as pd

# Import internal modules.
import xls_retrieve_data as xls

# Constants.
TISAX_FILE_NAME = 'C:/Users/jorge.silva/HUFGLOBAL/SGSI ISMS - ISMS Management - ISMS Management/I - TISAX Project/02 '\
                  '- Self assessment/05 - VDA 5_0_4 (05-01-2022)/VDA ISA 5.0.4_EN_archimate.xlsx'
TISAX_SHEET = 'Information Security'
REQUIREMENTS_FILE = 'C:/Users/jorge.silva/Downloads/tisax_high_requirements.xlsx'


def tisax_to_archi_high():

    # Create en empty output dataframe.
    df_data = pd.DataFrame(columns=['objectives', 'groups', 'high_requirements', 'high_sub-requirements'])
    df_objectives = pd.DataFrame(columns=['objectives'])
    df_groups = pd.DataFrame(columns=['groups'])
    df_high_reqs = pd.DataFrame(columns=['high_requirements'])
    df_high_subs = pd.DataFrame(columns=['high_sub-requirements'])

    # Get "high" requirements from TISAX check list.
    a_xls = xls.XlsRetrieveData(in_column=4, in_workbook=TISAX_FILE_NAME, in_worksheet=TISAX_SHEET, in_first_row=5)

    for a_n in range(5, a_xls.last_row + 1):

        # Example:
        # a_row =
        # ['1.1.1. The organization needs at least one information security policy. This reflects the importance and
        #          significance of information security and is adapted to the organization. Additional policies may be
        #          appropriate depending on the size and structure of the organization.',
        #   ['+ The requirements for information security have been determined and documented.',
        #    '  - The requirements are adapted to the goals of the organization.',
        #    "  - A policy has been created and approved by the organization's management.",
        #    '+ The policy includes objectives and the significance of information security within the organization.']
        # ]
        a_row = a_xls.get_high_row(a_n)

        if len(a_row) > 0:

            objective = a_row[0]
            objective_nr = objective.split(' ')[0]
            group = f'{objective_nr} High'
            requirement = ''
            requirement_nr = ''
            index_r = 0
            index_s = 0

            # Add objective and group to specific dataframes to write in different Excel sheets.
            df_obj = pd.DataFrame({'objectives': [objective]})
            df_objectives = pd.concat([df_objectives, df_obj], ignore_index=True)
            df_group = pd.DataFrame({'groups': [group]})
            df_groups = pd.concat([df_groups, df_group], ignore_index=True)

            for item in a_row[1]:

                item_list = item.split('+')  # if '+' means it is a requirement.
                if len(item_list) > 1:

                    index_r += 1
                    requirement_nr = f'{objective_nr}{index_r}'
                    index_s = 0
                    requirement = f'{requirement_nr}. {"".join(item_list[1:])}'

                    # Add requirement to dataframe to write in a different Excel sheet.
                    df_req = pd.DataFrame({'high_requirements': [requirement]})
                    df_high_reqs = pd.concat([df_high_reqs, df_req], ignore_index=True)

                item_list = item.split('-')  # if '-' means it is a sub-requirement.
                if len(item_list) > 1:

                    index_s += 1
                    sub_requirement = f'{requirement_nr}.{index_s}. {"".join(item_list[1:])}'

                    # Add sub-requirement to dataframe to write in a different Excel sheet.
                    df_sub_req = pd.DataFrame({'high_sub-requirements': [sub_requirement]})
                    df_high_subs = pd.concat([df_high_subs, df_sub_req], ignore_index=True)

                else:
                    sub_requirement = ''

                # Create a new "pandas" data frame to add to the output data frame.
                df_row_to_add = pd.DataFrame({'objectives': [objective], 'groups': [group],
                                              'high_requirements': [requirement],
                                              'high_sub-requirements': [sub_requirement]})
                # Add a new row (concatenate 2 data frames).
                df_data = pd.concat([df_data, df_row_to_add], ignore_index=True)

    # Write Excel files.
    a_excel_to_write = pd.ExcelWriter(REQUIREMENTS_FILE)
    df_data.to_excel(a_excel_to_write, sheet_name='all_requirements', index=False)
    df_objectives.to_excel(a_excel_to_write, sheet_name='objectives', index=False)
    df_groups.to_excel(a_excel_to_write, sheet_name='groups', index=False)
    df_high_reqs.to_excel(a_excel_to_write, sheet_name='high_requirements', index=False)
    df_high_subs.to_excel(a_excel_to_write, sheet_name='high_sub-requirements', index=False)
    a_excel_to_write.save()
    return


if __name__ == '__main__':
    tisax_to_archi_high()
