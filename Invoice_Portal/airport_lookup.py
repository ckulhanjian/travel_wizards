"""
airport_lookup.py - Static airport name lookup.

Maps truncated names from old-format invoices and IATA codes to full airport names.
No API calls. Runs instantly. Add entries as new airports appear in invoices.
"""

# ── IATA code -> (Full Airport Name, City) ────────────────────
# Covers major airports Travel Wizards clients would use.
# Add new codes as they appear in invoices.
IATA = {
    # North America - US
    "ANC": ("Ted Stevens Anchorage Intl", "Anchorage"),
    "ATL": ("Hartsfield-Jackson Atlanta Intl", "Atlanta"),
    "AUS": ("Austin-Bergstrom Intl", "Austin"),
    "BNA": ("Nashville Intl", "Nashville"),
    "BOS": ("Logan Intl", "Boston"),
    "BUR": ("Hollywood Burbank", "Burbank"),
    "BWI": ("Baltimore/Washington Intl", "Baltimore"),
    "CLE": ("Cleveland Hopkins Intl", "Cleveland"),
    "CLT": ("Charlotte Douglas Intl", "Charlotte"),
    "DCA": ("Ronald Reagan Washington National", "Washington DC"),
    "DEN": ("Denver Intl", "Denver"),
    "DFW": ("Dallas/Fort Worth Intl", "Dallas"),
    "DTW": ("Detroit Metropolitan Wayne County", "Detroit"),
    "EWR": ("Newark Liberty Intl", "Newark"),
    "FLL": ("Fort Lauderdale-Hollywood Intl", "Fort Lauderdale"),
    "HNL": ("Daniel K. Inouye Intl", "Honolulu"),
    "IAD": ("Washington Dulles Intl", "Washington DC"),
    "IAH": ("George Bush Intercontinental", "Houston"),
    "IND": ("Indianapolis Intl", "Indianapolis"),
    "JFK": ("John F. Kennedy Intl", "New York"),
    "LAS": ("Harry Reid Intl", "Las Vegas"),
    "LAX": ("Los Angeles Intl", "Los Angeles"),
    "LGA": ("LaGuardia", "New York"),
    "MCI": ("Kansas City Intl", "Kansas City"),
    "MCO": ("Orlando Intl", "Orlando"),
    "MDW": ("Chicago Midway Intl", "Chicago"),
    "MIA": ("Miami Intl", "Miami"),
    "MRS": ("Marseille Provence Airport", "Marignane"),
    "MSP": ("Minneapolis-St Paul Intl", "Minneapolis"),
    "MSY": ("Louis Armstrong New Orleans Intl", "New Orleans"),
    "OAK": ("Oakland Intl", "Oakland"),
    "ONT": ("Ontario Intl", "Ontario"),
    "ORD": ("O'Hare Intl", "Chicago"),
    "PBI": ("Palm Beach Intl", "West Palm Beach"),
    "PDX": ("Portland Intl", "Portland"),
    "PHL": ("Philadelphia Intl", "Philadelphia"),
    "PHX": ("Phoenix Sky Harbor Intl", "Phoenix"),
    "PIT": ("Pittsburgh Intl", "Pittsburgh"),
    "RDU": ("Raleigh-Durham Intl", "Raleigh"),
    "RSW": ("Southwest Florida Intl", "Fort Myers"),
    "SAN": ("San Diego Intl", "San Diego"),
    "SAT": ("San Antonio Intl", "San Antonio"),
    "SEA": ("Seattle-Tacoma Intl", "Seattle"),
    "SFO": ("San Francisco Intl", "San Francisco"),
    "SJC": ("San Jose Intl", "San Jose"),
    "SLC": ("Salt Lake City Intl", "Salt Lake City"),
    "SMF": ("Sacramento Intl", "Sacramento"),
    "SNA": ("John Wayne/Orange County", "Santa Ana"),
    "SRQ": ("Sarasota Bradenton Intl", "Sarasota"),
    "STL": ("St. Louis Lambert Intl", "St. Louis"),
    "TPA": ("Tampa Intl", "Tampa"),

    # North America - Canada
    "YEG": ("Edmonton Intl", "Edmonton"),
    "YOW": ("Ottawa Macdonald-Cartier Intl", "Ottawa"),
    "YUL": ("Montreal-Trudeau Intl", "Montreal"),
    "YVR": ("Vancouver Intl", "Vancouver"),
    "YYC": ("Calgary Intl", "Calgary"),
    "YYZ": ("Toronto Pearson Intl", "Toronto"),

    # North America - Mexico / Caribbean
    "CUN": ("Cancun Intl", "Cancun"),
    "GDL": ("Guadalajara Intl", "Guadalajara"),
    "MEX": ("Mexico City Intl", "Mexico City"),
    "MBJ": ("Sangster Intl", "Montego Bay"),
    "NAS": ("Lynden Pindling Intl", "Nassau"),
    "PUJ": ("Punta Cana Intl", "Punta Cana"),
    "PVR": ("Gustavo Diaz Ordaz Intl", "Puerto Vallarta"),
    "SJD": ("Los Cabos Intl", "San Jose del Cabo"),
    "SJU": ("Luis Munoz Marin Intl", "San Juan"),

    # Europe
    "AMS": ("Amsterdam Schiphol", "Amsterdam"),
    "ATH": ("Athens Intl", "Athens"),
    "BCN": ("Barcelona-El Prat", "Barcelona"),
    "BRU": ("Brussels Airport", "Brussels"),
    "CDG": ("Charles de Gaulle", "Paris"),
    "CPH": ("Copenhagen Airport", "Copenhagen"),
    "DUB": ("Dublin Airport", "Dublin"),
    "DUS": ("Dusseldorf Airport", "Dusseldorf"),
    "EDI": ("Edinburgh Airport", "Edinburgh"),
    "FCO": ("Leonardo da Vinci-Fiumicino", "Rome"),
    "FRA": ("Frankfurt Airport", "Frankfurt"),
    "GVA": ("Geneva Airport", "Geneva"),
    "HEL": ("Helsinki-Vantaa", "Helsinki"),
    "IST": ("Istanbul Airport", "Istanbul"),
    "LGW": ("London Gatwick", "London"),
    "LHR": ("London Heathrow", "London"),
    "LIS": ("Humberto Delgado Airport", "Lisbon"),
    "MAD": ("Adolfo Suarez Madrid-Barajas", "Madrid"),
    "MAN": ("Manchester Airport", "Manchester"),
    "MUC": ("Munich Airport", "Munich"),
    "MXP": ("Milan Malpensa", "Milan"),
    "NCE": ("Nice Cote d'Azur", "Nice"),
    "ORY": ("Paris Orly", "Paris"),
    "OSL": ("Oslo Gardermoen", "Oslo"),
    "PRG": ("Vaclav Havel Airport", "Prague"),
    "STN": ("London Stansted", "London"),
    "VCE": ("Venice Marco Polo", "Venice"),
    "VIE": ("Vienna Intl", "Vienna"),
    "ZRH": ("Zurich Airport", "Zurich"),

    # Asia-Pacific
    "BKK": ("Suvarnabhumi Airport", "Bangkok"),
    "DEL": ("Indira Gandhi Intl", "Delhi"),
    "HKG": ("Hong Kong Intl", "Hong Kong"),
    "HND": ("Tokyo Haneda", "Tokyo"),
    "ICN": ("Incheon Intl", "Seoul"),
    "KIX": ("Kansai Intl", "Osaka"),
    "MNL": ("Ninoy Aquino Intl", "Manila"),
    "NRT": ("Narita Intl", "Tokyo"),
    "PEK": ("Beijing Capital Intl", "Beijing"),
    "PVG": ("Shanghai Pudong Intl", "Shanghai"),
    "SGN": ("Tan Son Nhat Intl", "Ho Chi Minh City"),
    "SIN": ("Singapore Changi", "Singapore"),
    "SYD": ("Sydney Kingsford Smith", "Sydney"),
    "TPE": ("Taiwan Taoyuan Intl", "Taipei"),

    # South America
    "BOG": ("El Dorado Intl", "Bogota"),
    "EZE": ("Ministro Pistarini Intl", "Buenos Aires"),
    "GIG": ("Rio de Janeiro-Galeao", "Rio de Janeiro"),
    "GRU": ("Sao Paulo-Guarulhos", "Sao Paulo"),
    "LIM": ("Jorge Chavez Intl", "Lima"),
    "SCL": ("Santiago Intl", "Santiago"),

    # Middle East / Africa
    "ADD": ("Addis Ababa Bole Intl", "Addis Ababa"),
    "AMM": ("Queen Alia Intl", "Amman"),
    "CAI": ("Cairo Intl", "Cairo"),
    "CMN": ("Mohammed V Intl", "Casablanca"),
    "CPT": ("Cape Town Intl", "Cape Town"),
    "DOH": ("Hamad Intl", "Doha"),
    "DXB": ("Dubai Intl", "Dubai"),
    "JNB": ("O.R. Tambo Intl", "Johannesburg"),
    "NBO": ("Jomo Kenyatta Intl", "Nairobi"),
    "TLV": ("Ben Gurion Intl", "Tel Aviv"),

    # Oceania
    "AKL": ("Auckland Airport", "Auckland"),
    "MEL": ("Melbourne Tullamarine", "Melbourne"),
    "PPT": ("Tahiti Faa'a Intl", "Papeete"),

    "FLR": ("Peretola Airport", "Florence"),
    "KOA": ("Ellison Onizuka Kona Intl", "Kona")
}

# ── Truncated name -> IATA code ───────────────────────────────
# Maps the truncated city names that appear in old-format invoices.
# Key is UPPERCASE, matched against the raw parsed city name.
TRUNCATED = {
    # Exact matches (no truncation)
    "DENVER": "DEN",
    "MARSEILLE": "MRS",
    "VANCOUVER": "YVR",
    "ANCHORAGE": "ANC",
    "BARCELONA": "BCN",
    "LISBON": "LIS",
    "LONDON": "LHR",
    "PARIS": "CDG",
    "ROME": "FCO",
    "TOKYO": "NRT",
    "AMSTERDAM": "AMS",
    "DUBLIN": "DUB",
    "FRANKFURT": "FRA",
    "MUNICH": "MUC",
    "VENICE": "VCE",
    "ZURICH": "ZRH",
    "HONG KONG": "HKG",
    "SINGAPORE": "SIN",
    "SYDNEY": "SYD",
    "MIAMI": "MIA",
    "BOSTON": "BOS",
    "ATLANTA": "ATL",
    "ORLANDO": "MCO",
    "DALLAS": "DFW",
    "HOUSTON": "IAH",
    "PHOENIX": "PHX",
    "PORTLAND": "PDX",
    "HONOLULU": "HNL",
    "NASHVILLE": "BNA",
    "CANCUN": "CUN",
    "MONTREAL": "YUL",
    "TORONTO": "YYZ",
    "EDMONTON": "YEG",
    "CALGARY": "YYC",
    "OTTAWA": "YOW",

    # Truncated names from old-format invoices
    "SAN FRANCISCO/SAN FRANC": "SFO",
    "SAN FRANCISCO": "SFO",
    "SEATTLE/TACOMA INTERNAT": "SEA",
    "SEATTLE/TACOMA": "SEA",
    "SEATTLE": "SEA",
    "TORONTO/LESTER B PEARSO": "YYZ",
    "TORONTO/LESTER B PEARSON": "YYZ",
    "LOS ANGELES/LOS ANGELE": "LAX",
    "LOS ANGELES": "LAX",
    "NEW YORK/JOHN F KENNED": "JFK",
    "NEW YORK/JOHN F KENNEDY": "JFK",
    "NEW YORK/NEWARK": "EWR",
    "NEW YORK/LAGUARDIA": "LGA",
    "WASHINGTON/DULLES INTER": "IAD",
    "WASHINGTON/DULLES": "IAD",
    "WASHINGTON/NATIONAL": "DCA",
    "CHICAGO/O'HARE INTERNA": "ORD",
    "CHICAGO/O'HARE": "ORD",
    "CHICAGO/OHARE": "ORD",
    "CHICAGO": "ORD",
    "DALLAS/FORT WORTH INTE": "DFW",
    "DALLAS/FORT WORTH": "DFW",
    "HOUSTON/GEORGE BUSH IN": "IAH",
    "HOUSTON/GEORGE BUSH": "IAH",
    "LAS VEGAS": "LAS",
    "SAN DIEGO": "SAN",
    "SAN JOSE": "SJC",
    "SALT LAKE CITY": "SLC",
    "MINNEAPOLIS/ST PAUL IN": "MSP",
    "MINNEAPOLIS": "MSP",
    "DETROIT/METROPOLITAN W": "DTW",
    "DETROIT": "DTW",
    "CHARLOTTE": "CLT",
    "PHILADELPHIA": "PHL",
    "FORT LAUDERDALE": "FLL",
    "TAMPA": "TPA",
    "NEW ORLEANS": "MSY",
    "PARIS/CHARLES DE GAULL": "CDG",
    "PARIS/CHARLES DE GAULLE": "CDG",
    "PARIS/ORLY": "ORY",
    "LONDON/HEATHROW": "LHR",
    "LONDON/GATWICK": "LGW",
    "LONDON/STANSTED": "STN",
    "ROME/FIUMICINO": "FCO",
    "MILAN/MALPENSA": "MXP",
    "MADRID": "MAD",
    "ISTANBUL": "IST",
    "DUBAI": "DXB",
    "DOHA": "DOH",
    "MEXICO CITY": "MEX",
    "SAO PAULO": "GRU",
    "BUENOS AIRES": "EZE",
    "BANGKOK": "BKK",
    "SEOUL/INCHEON": "ICN",
    "SARASOTA": "SRQ",
    "SACRAMENTO": "SMF",
    "OAKLAND": "OAK",
    "SAN ANTONIO": "SAT",
    "FORT MYERS": "RSW",
    "WEST PALM BEACH": "PBI",
    "KANSAS CITY": "MCI",
    "CLEVELAND": "CLE",
    "INDIANAPOLIS": "IND",
    "AUSTIN": "AUS",
    "ST LOUIS": "STL",
    "RALEIGH": "RDU",
    "PITTSBURGH": "PIT",
    "BURBANK": "BUR",
    "FLORENCE": "FLR",
    "KONA": "KOA",
    "KONA/KAILUA": "KOA"
}


def lookup_airport(city_name: str) -> dict:
    """
    Look up a city name (possibly truncated) and return full airport info.

    Returns: {
        "iata": "SFO",
        "airport": "San Francisco Intl",
        "city": "San Francisco",
        "display": "San Francisco Intl (SFO)"
    }
    or None if not found (falls back to original name).
    """
    if not city_name:
        return None

    key = city_name.strip().upper()

    # Try exact match in truncated map
    if key in TRUNCATED:
        code = TRUNCATED[key]
        if code in IATA:
            airport, city = IATA[code]
            return {
                "iata": code,
                "airport": airport,
                "city": city,
                "display": f"{airport}, {city} ({code})",
            }

    # Try prefix match (for unknown truncations)
    for trunc_name, code in TRUNCATED.items():
        if key.startswith(trunc_name[:8]) and code in IATA:
            airport, city = IATA[code]
            return {
                "iata": code,
                "airport": airport,
                "city": city,
                "display": f"{airport}, {city} ({code})",
            }

    return None


def resolve_city(city_name: str) -> str:
    """
    Return a clean display string for a city name.
    If found in lookup: "San Francisco Intl, San Francisco (SFO)"
    If not found: title-cased original name.
    """
    result = lookup_airport(city_name)
    if result:
        return result["display"]
    # Clean up the raw name
    return " ".join(w.capitalize() for w in city_name.lower().split("/")[0].split())


if __name__ == "__main__":
    # Test with known truncated names
    tests = [
        "DENVER",
        "VANCOUVER",
        "ANCHORAGE",
        "SEATTLE/TACOMA INTERNAT",
        "SAN FRANCISCO/SAN FRANC",
        "TORONTO/LESTER B PEARSO",
        "BARCELONA",
        "LISBON",
        "UNKNOWN CITY",
    ]
    for t in tests:
        print(f"  {t:30} -> {resolve_city(t)}")
