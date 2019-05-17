import datetime
import json
import ssl
from typing import Tuple

from librouteros import connect
from librouteros.login import login_plain, login_token
import tablib

# config
routeros_63_or_higher = False
username = 'api'
password = 'Auo7xmBJ8DgoWFH7'
host = "191.232.243.34"
port = 8729

# ssl context
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# for new (plain text password)
if routeros_63_or_higher:
    method = (login_plain, )
else:
    # for old (with token)
    method = (login_token, )

api = connect(
    username=username,
    password=password,
    host=host,
    ssl_wrapper=ctx.wrap_socket,
    login_methods=method,
    port=port
)


def tuple_to_dataset(input_tuple: Tuple, title: str, db):
    # skip empty tables
    if len(input_tuple) != 0:
        # ugly hack to fix invalid dimensions error for the files sheet
        if title == 'files':
            for dict_item in input_tuple:
                if 'size' not in dict_item:
                    dict_item['size'] = ' '
                if 'contents' not in dict_item:
                    dict_item['contents'] = ' '
        if title == 'users':
            for dict_item in input_tuple:
                if 'comment' not in dict_item:
                    dict_item['comment'] = ' '
        json_string = json.dumps(list(input_tuple))
        dataset = tablib.Dataset(title=title)
        dataset.json = json_string
        db.add_sheet(dataset)


try:
    logs = api(cmd='/log/print')
    dns_cache = api(cmd='/ip/dns/cache/print')
    dns_static = api(cmd='/ip/dns/static/print')
    dhcp_server = api(cmd='/ip/dhcp-server/print')
    dhcp_relay = api(cmd="/ip/dhcp-relay/print")
    dhcp_client = api(cmd='/ip/dhcp-client/print')
    users = api(cmd='/user/print')
    arp = api(cmd='/ip/arp/print')
    files = api(cmd='/file/print')
    ip_route = api(cmd='/ip/route/print')
    bgp_ads = api(cmd="/routing/bgp/advertisements/print")

    db = tablib.Databook()

    tuple_to_dataset(logs, 'logs', db)
    tuple_to_dataset(dns_cache, 'dns_cache', db)
    tuple_to_dataset(dhcp_client, 'dhcp_client', db)
    tuple_to_dataset(dhcp_relay, 'dhcp_relay', db)
    tuple_to_dataset(dhcp_server, 'dhcp_server', db)
    tuple_to_dataset(users, 'users', db)
    tuple_to_dataset(arp, 'arp', db)
    tuple_to_dataset(files, 'files', db)
    tuple_to_dataset(ip_route, 'ip_route', db)
    tuple_to_dataset(bgp_ads, 'bgp_advertisements', db)

    date = str(datetime.datetime.now())
    open('logs/logs-' + date + '.xlsx', 'wb').write(db.xlsx)
except Exception as e:
    print('Exception occured.')
    print(e)
