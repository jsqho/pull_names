import requests

def get_json_result(doi):
	API_URL = "http://api.crossref.org/works/"
	r = requests.get(API_URL + "doi")
	if r.status_code == "200":
		print r.json()
	else:
		print "Page " + r.status_code