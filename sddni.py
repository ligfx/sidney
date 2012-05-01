#!/usr/bin/env python

from contextlib import closing
import re
import sqlite3
import sys
import xlrd

print "Reading Student Directory..."
xls = xlrd.open_workbook("Student Directory Spring 2012.xls")
sheet = xls.sheet_by_name("pcregex (13)")
column_names = {
	"name": 0,
	"phone": 2,
	"gender": 8,
	"year": 10,
	"building": 12,
	"room": 14,
	"mailbox": 16
}
def cell_to_string(cell):
	if (cell.ctype == 2):
		return str(int(cell.value))
	else:
		return str(cell.value)
data = dict((key, [cell_to_string(cell) for cell in sheet.col(val)[3:]]) for key, val in column_names.iteritems())

building_names = [
	["WIG", "Wig"],
	["LYON", "Lyon"],
	["OLD", "Oldenborg"],
	["SMI", "Smiley"],
	["MUDD", "BLSD", "Mudd Blaisdell"],
	["N-CL", "Norton", "Clark III"],
	["CL-I", "Clark I"],
	["OFFC"],
	["GIBS"],
	["EXCH-PTZ"],
	["SNTG"],
	["EXCH-CMC"],
	["POM"],
	["CL-V", "Clark V"],
	["HAR", "Harwood"],
	["C241"],
	["WALK", "Walker"],
	["LWRY", "Lawry"],
	["EXCH-SCR"],
]
print "Creating student information data store..."
db = sqlite3.connect(":memory:")
with closing(db.cursor()) as c:
	c.execute('''CREATE TABLE 'people' (id integer primary key, name text, phone text, gender text, year text, building text, room text);''')
	for i in range(len(data["name"])):
		val = [data[key][i] for key in ("name", "phone", "gender", "year", "building", "room")]
		c.execute('''INSERT INTO 'people' values (null, ?, ?, ?, ?, ?, ?);''', val)
	c.execute('''CREATE TABLE 'building_aliases' (id integer primary key, building_id integer, alias text);''');
	for (id, building) in zip(range(len(building_names)), building_names):
		for alias in building:
			c.execute('''INSERT INTO 'building_aliases' values (null, ?, ?);''', (id, alias))
	db.commit()
print "Done!"

class PersonRepository:
	fields = ["id", "name", "phone", "gender", "year", "building", "room"]
	def __init__(self, db):
		self.db = db
	def search_by_name(self, keywords):
		c = self.db.cursor()
		condition = " AND ".join(["( 'people'.'name' LIKE ? )"] * len(keywords))
		return map(
			lambda p: Person(self.tuple_to_dict(p)),
			c.execute('''
				SELECT * FROM 'people' WHERE (%s)
				''' % condition, ['%' + k + '%' for k in keywords]
				).fetchall()
			)
	def tuple_to_dict(self, tuple):
		return dict((key, value) for (key, value) in zip(self.fields, tuple))

class BuildingAliasRepository:
	fields = ["id", "building_id", "alias"]
	def __init__(self, db):
		self.db = db
	def find_aliases_for_name(self, name):
		c = self.db.cursor()
		return map(
			lambda b: BuildingAlias(self.tuple_to_dict(b)),
			c.execute('''
				SELECT * FROM building_aliases WHERE building_id IN (
					SELECT building_id FROM building_aliases WHERE alias = ?
				)''', [name]
			).fetchall()
		)
	def tuple_to_dict(self, tuple):
		return dict((key, value) for key, value in zip(self.fields, tuple))

class Person:
	def __init__(self, opts={}):
		self.name = opts["name"]
		self.building = opts["building"]
		self.room = opts["room"]
		self.phone = opts["phone"]
	def __repr__(self):
		return "<Person: %s>" % self.name

class BuildingAlias:
	def __init__(self, opts={}):
		self.building_id = opts["building_id"]
		self.alias = opts["alias"]
	def __repr__(self):
		return "<BuildingAlias: %s>" % self.alias

def room_to_number(room_name):
	m = re.match("(\d+)", room_name)
	if m:
		return m.group()
	else:
		return None

def room_names_match(a, b):
	if a == b: return True
	if a == room_to_number(b): return True
	if room_to_number(a) == b: return True
	return False

def indexitems(enum):
	return zip(range(len(enum)), enum)

class DormNetworkInventory:
	def __init__(self):
		self.xls = xlrd.open_workbook("Dorm Network Inventory.xls")
	def find_all_by_names(self, names):
		return [(a, DormNetworkInventoryBuilding(self.xls.sheet_by_name(a))) for a in names if a in self.xls.sheet_names()]

class DormNetworkInventoryBuilding:
	def __init__(self, sheet):
		self.plates = [cell_to_string(cell) for cell in sheet.col(0)[3:]]
		self.ports = [cell_to_string(cell) for cell in sheet.col(1)[3:]]
		self.rooms = [cell_to_string(cell) for cell in sheet.col(2)[3:]]
	def find_rooms_by_name(self, name):
		return [(i,room_name) for (i, room_name) in indexitems(self.rooms) if room_names_match(name, room_name)]
	def find_jacks_by_rooms(self, rooms):
		return map(
			lambda (index, room_name): Jack({
				"plate": self.plates[index],
				"port": self.ports[index],
				"room": room_name
			}),
			rooms
		)

class Jack:
	def __init__(self, opts={}):
		self.plate = opts["plate"]
		self.port = opts["port"]
		self.room = opts["room"]

class PersonResult:
	def __init__(self):
		self.person = None
		self.building_aliases = []
		self.buildings = []
	def add_building(self, alias, jacks):
		self.buildings.append({
			"alias": alias,
			"jacks": jacks
		})

class PersonResultsContext:
	def __init__(self, db, keywords):
		self.results = []

		person_repository = PersonRepository(db)
		building_alias_repository = BuildingAliasRepository(db)
		dorm_network_inventory = DormNetworkInventory()
		people = person_repository.search_by_name(keywords)
		for person in people:
			result = PersonResult()
			result.person = person
			result.building_aliases = building_alias_repository.find_aliases_for_name(person.building)
			for (alias, building) in dorm_network_inventory.find_all_by_names(b.alias for b in result.building_aliases):
				rooms = building.find_rooms_by_name(person.room)
				jacks = building.find_jacks_by_rooms(rooms)
				result.add_building(alias, jacks)
			self.results.append(result)

def find_person(keywords):
	results = PersonResultsContext(db, keywords).results

	print "Searching for people who match '%s'..." % (" ".join(keywords))
	print
	if len(results) == 0:
		print "Couldn't find anyone matching '%s' :(" % (" ".join(keywords))
	for result in results:
		person = result.person
		print "Found person '%s' who lives in '%s' room '%s'!" % (person.name, person.building, person.room)
		print "Phone: %s" % person.phone
		print "Possible aliases for building '%s' include:" % person.building
 		for b in result.building_aliases:
 			print " * '%s'" % b.alias
 		print "Searching for Dorm Network Inventory data..."
 		if len(result.buildings) == 0:
 			print "Couldn't find Dorm Network Inventory data for building '%s' :(" % person.building
 		for b in result.buildings:
 			print "Found Dorm Network Inventory data for alias '%s'!" % b["alias"]
 			print "Searching for room '%s' in '%s'..." % (person.room, b["alias"])
			if len(b["jacks"]) > 0:
				print "Possible port numbers for this person:"
				for j in b["jacks"]:
					print " * '%s%s' in room '%s' in building '%s' for person '%s'" % (j.plate, j.port, j.room, b["alias"], person.name)
			else:
				print "Couldn't find room '%s' in '%s' :(" % (person.room, b["alias"])
		print

print
print "Welcome to the automated Student Directory Dorm Network Inventory (SDDNI) Cross-Referencer!"
print "You are running SDDNI (Sidney) v0.1 the Amorous Armadillo, (c) Michael Maltese 2012"
print "All information is the property of Pomona College."
print
sys.stdout.write("Who are you looking for? ")
keywords = re.split("\s+", sys.stdin.readline().strip("\n"))
if keywords == ['']:
	print "Can't search for an empty name!"
	print
	sys.exit()
find_person(keywords)