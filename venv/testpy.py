import tableauserverclient as TSC

str='泛微-群组-致同群组'
print(str[:-5])
# sign in, etc

tableau_auth = TSC.TableauAuth('GT_Account', 'GT2020123456', site_id='')
server = TSC.Server('https://fas-gt.com:8000',use_server_version=True)



  # query the sites
with server.auth.sign_in(tableau_auth):
    # all_sites, pagination_item = server.sites.get()
    #
    # # print all the site names and ids
    # for site in all_sites:
    #     print(site.id, site.name, site.content_url, site.state)

    all_workbooks,pagination_item = server.workbooks.get()
    for wbks in all_workbooks:
        print(wbks.id,wbks.name)

    all_users, pagination_item = server.users.get()
    print("\nThere are {} user on site: ".format(pagination_item.total_available))
    print([user.name for user in all_users])
    print([user.id for user in all_users])
    #
    # all_workbooks_items, pagination_item = server.workbooks.get()
    # # print names of first 100 workbooks
    # print([workbook.name for workbook in all_workbooks_items])
    # print([workbook.id for workbook in all_workbooks_items])
