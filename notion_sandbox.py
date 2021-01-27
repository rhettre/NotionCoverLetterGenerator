from notion.client import NotionClient

my_token = "3ca4924184f463dc694abc3700f58f28028c42ffd83d9d5be480121eb4c1ad864f1c3604257b1879b1d88d455f09ae93197137d9268801d75bf5508f947a53b2b92bdbe745f249eb5d9d2a7dd895"
client = NotionClient(token_v2 = my_token)

#lookup ids for different columns
position_id = 'xyVl'
status_id = 'RpsQ'
location_id = 'SXHt'
where_did_you_find_job_id = 'Hv_n'
why_this_job_id = 'WB`~'
name_id = 'title'

cv = client.get_collection_view("https://www.notion.so/rhettre/a1808aae40504e88ba013ccf46069a1a?v=bb1cc2c7f7d64d3298ed4068c7df1ca5")

rows = cv.collection.get_rows()

first_row = rows[0]
# print(cv.collection.get_schema_properties())
# print(first_row.get_property(name_id))
# print(first_row.get_property(position_id))
print(first_row.get_property(status_id))
# print(first_row.get_property(location_id))
# print(first_row.get_property(where_did_you_find_job_id))
# print(first_row.get_property(why_this_job_id))


for row in rows:
    #Collect all where the the cover letter ID is false 
    if row.get_property(status_id) == "Information Complete":

        #Start Parsing the Data from the Row into your cover letter verbiage
        #Export to Microsoft Word file using python-docx 
        #Save as PDF

        #Set the cover letter ID to true

