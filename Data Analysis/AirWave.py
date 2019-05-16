# -*- coding: utf-8 -*-


from airwaveapiclient import AirWaveAPIClient
from airwaveapiclient import APList
from airwaveapiclient import APDetail
from pprint import pprint
import re
from collections import OrderedDict


def main():
    """Sample main."""
    #################################################
    # Settings ######################################
    #################################################

    airwaveServerUrls = [ 'https://amp-comp-01.net.ku.edu', 'https://amp-comp-02.net.ku.edu',
                          'https://amp-ellx-01.net.ku.edu' ]
    username = ''
    password = ''

    nodes = {}

    # Login to each airwave server and build a list
    for url in airwaveServerUrls:
        airwave = AirWaveAPIClient(username=username,
                                   password=password,
                                   url=url)
        #print("Logging into " + url + " with username '" + username + "'...")
        airwave.login()

        # collect APs -- only with two letter code "WA"
        #print("Obtaining AP information...")
        res = airwave.ap_list()
        p = re.compile('^\w+\-WA', re.IGNORECASE)
        if res.status_code == 200:
            xml = res.text
            ap_list = APList(xml)
            for ap_node in ap_list:
                #pprint(ap_node)
                if p.match(ap_node['name']):
                    radio_macs = [ ]
                    for ap_radio in ap_node['radio']:
                        if isinstance(ap_radio, OrderedDict):
                            radio_macs.append(str(ap_radio['radio_mac']))

                    nodes[ap_node['name']] = { 'lan_ip' : ap_node['lan_ip'], 'mac_radios' : radio_macs }
        # logout
        airwave.logout()

    for key in sorted(nodes.keys()):
        line = key
        if nodes[key]['lan_ip']:
            line += ',' + nodes[key]['lan_ip']
        else:
            line += ','

        if nodes[key]['mac_radios']:
            for radio in nodes[key]['mac_radios']:
                line += ',' + radio
        else:
            line += ','

        print(line)


    # Example for more detail:
    #for ap_node in ap_list:
    #    res = airwave.ap_detail(ap_node['@id'])
    #    if res.status_code == 200:
    #        xml = res.text
    #        ap_detail = APDetail(xml)
    #        pprint(ap_detail)


if __name__ == "__main__":

    main()
