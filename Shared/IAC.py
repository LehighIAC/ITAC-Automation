"""
(Purpose) IAC.py is a module that contains functions used in the IAC report
"""


def ez_degree_days(ZIP: str, mode: str, temp: int) -> float:
    """
    Automatically calculate degree days based on daily average temperature
    The result should be equal to degreedays.net
    :param ZIP: ZIP code as string
    :param mode: "heating" or "cooling" as string
    :param temp: Base temperature as integer
    :return: Degree days as float
    """
    from meteostat import Stations, Daily, units
    from datetime import datetime
    import pgeocode

    # Get coordinate from ZIP code
    location = pgeocode.Nominatim('us').query_postal_code(ZIP)

    # Get closest weather station
    station = Stations().nearby(location.latitude, location.longitude).fetch(1).index[0]

    # 4 years of data, by default
    gap_year = 4
    start = datetime(datetime.now().year - gap_year, 1, 1)
    end = datetime(datetime.now().year - 1, 12, 31)

    # Get hourly data
    data = Daily(station, start, end)
    data = data.convert(units.imperial)
    data = data.normalize()
    # https://github.com/meteostat/meteostat-python/issues/130
    #data = data.interpolate()
    data = data.fetch()

    if mode == "cooling":
        # Calculate cooling degree days
        data['dd'] = data.apply(lambda x: max((x['tavg'] - temp), 0), axis=1)
        dd = data.dd.sum() / gap_year
    elif mode == "heating":
        # Calculate heating degree days
        data['dd'] = data.apply(lambda x: max((temp - x['tavg']), 0), axis=1)
        dd = data.dd.sum() / gap_year
    else:
        print("Mode must be 'heating' or 'cooling'")
        exit()
    return dd

def degree_days(ZIP: str, mode: str, basetemp: int, setback: int=None, hours: list=[9,17], weekend: list=[]) -> float:
    """
    Automatically calculate degree days based on hourly data
    The result is usually higher than degreedays.net
    :param ZIP: ZIP code as string
    :param mode: "heating" or "cooling" as string
    :param basetemp: Base temperature as integer
    :param setback: Setback temperature as integer, default is None
    :param hours: Operating hours as list of integer, default is [9, 17] (9am to 5pm)
    :param weekend: List of weekend schedule as list, default is []. Example: [5,6] (Saturday and Sunday)
    :return: Degree days as float
    """
    from meteostat import Stations, Hourly, units
    from datetime import datetime
    import pgeocode

    # Get coordinate from ZIP code
    location = pgeocode.Nominatim('us').query_postal_code(ZIP)

    # Get closest weather station
    station = Stations().nearby(location.latitude, location.longitude).fetch(1).index[0]

    # 4 years of data, by default
    gap_year = 1
    starttime = datetime(datetime.now().year - gap_year, 1, 1)
    endtime = datetime(datetime.now().year - 1, 12, 31, 23, 59)

    # Get hourly data
    data = Hourly(station, starttime, endtime)
    data = data.convert(units.imperial)
    data = data.normalize()
    # https://github.com/meteostat/meteostat-python/issues/130
    #data = data.interpolate()
    data = data.fetch()

    # Add a column for set temperature and set it to base temperature
    data['Tset'] = basetemp

    # If setback temperature is provided
    if setback != None:
        # Override time outside weekday hours
        data['hour'] = data.index.hour
        data['Tset'] = data.apply(lambda x: setback if (x['hour'] <= hours[0] or x['hour'] >= hours[1]) else x['Tset'], axis=1)
        # Override time on weekend
        for holiday in weekend:
            data['day'] = data.index.weekday
            data['Tset'] = data.apply(lambda x: setback if (x['day'] == holiday) else x['Tset'], axis=1)

    if mode == "cooling":
        # Calculate cooling degree days
        data['dd'] = data.apply(lambda x: max((x['temp'] - x['Tset']), 0), axis=1)
        dd = data.dd.sum() / gap_year / 24
    elif mode == "heating":
        # Calculate heating degree days
        data['dd'] = data.apply(lambda x: max((x['Tset'] - x['temp']), 0), axis=1)
        dd = data.dd.sum() / gap_year / 24
    else:
        print("Mode must be 'heating' or 'cooling'")
        exit()
    return dd

def validate_arc(ARC):
    """
    Validate ARC input
    :param ARC: Full ARC number as a string
    """
    # json5 is too slow, use json instead.
    import os, json
    # Validate if ARC is in x.xxxx.xxx format
    ARCsplit = ARC.split('.')
    if len(ARCsplit) != 3:
        print("ARC number must be in x.xxxx.x format")
        exit()
    # if ARC split are nut full numbers
    for i in range(len(ARCsplit)):
        if ARCsplit[i].isdigit() == False:
            print("ARC number must be in x.xxxx.x format")
            exit()
    
    # Parse ARC code
    code = ARCsplit[0] + '.' + ARCsplit[1]
    # Read ARC.json5 as dictionary
    arc_path = os.path.dirname(os.path.abspath(__file__))
    ARCdict = json.load(open(os.path.join(arc_path, 'ARC.json')))
    desc = ARCdict[code]
    if desc == None:
        print("ARC code not found.")
    else:
        print(code + ": "+ desc)

    # Parse application code
    app = ARCsplit[2]
    if app == '1':
        print("Application code 1: Manufacturing Process")
    elif app == '2':
        print("Application code 2: Process Support")
    elif app == '3':
        print("Application code 3: Building and Grounds")
    elif app == '4':
        print("Application code 4: Administrative")
    else:
        print("Application code not found.")

    print("")

def grouping_num(dic: dict) -> dict:
    """
    Add thousand separator to numbers in a dictionary and format it to string
    :param dic: EasyDict
    :return: Dictionary with keys in thousand separator
    """
    import locale
    # set locale to US
    locale.setlocale(locale.LC_ALL, 'en_US')
    for key in dic.keys():
        if type(dic[key]) == int:
            dic[key] = locale.format_string('%d', dic[key], grouping=True)
        elif type(dic[key]) == float:
            dic[key] = locale.format_string('%g', dic[key], grouping=True)
        else:
            pass
    return dic

def dollar(varlist: list, dic: dict, digits: int=0) -> str:
    """
    Format numbers in a dictionary and to currency string
    :param varlist: List of keys in the dictionary
    :param dic: EasyDict
    :param digits: Number of digits, default is 0
    :return: Dictionary with keys in formatted currency string
    """
    import locale
    locale.setlocale(locale.LC_ALL, 'en_US')
    locale._override_localeconv={'frac_digits':digits}
    for var in varlist:
        dic[var] = locale.currency(dic[var], grouping=True)
    return dic

def combine_words(words: list) -> str:
    """
    :param words: list of strings
    :return: string of words separated by "," and "and"
    """
    combined = ""
    for i in range(len(words)):  
        combined = combined + words[i]    
        if i < len(words) - 2:
            combined = combined + ', '
        if i == len(words) - 2:
            combined = combined + ' and ' 
        else:
            pass
    return combined

def add_image(doc, tag: str, image_path: str, wd):
    """
    Add image to Word document, search for tag in doc and replace with the image
    :param doc: Document
    :param tag: Image tag string
    :param image_path: Path to the image
    :param wd: Image width
    :return: None
    """
    found_tag = False
    for p in doc.paragraphs:
        if tag in p.text:
            p.text = p.text.replace(tag, '')
            r = p.add_run()
            r.add_picture(image_path, width=wd)
            found_tag = True
            break
    if found_tag == False:
        print("Tag "+ tag +" not found")

def add_eqn(doc, tag: str, eqn_input):
    """
    Add equation to Word document, search for eqn in doc and replace with eqn_input
    :param doc: Document
    :param tag: Equation tag string
    :param eqn_input: Equation input
    :return: None
    """
    found_tag = False
    for p in doc.paragraphs:
        if tag in p.text:
            p.text = p.text.replace(tag, '')
            word_math = latex2word(eqn_input)
            p._element.append(word_math)
            found_tag = True
            break
    if found_tag == False:
        print("Tag "+ tag +" not found")

def latex2word(latex_input):
    """
    Convert LaTeX equation to Word equation
    :param latex_input: LaTeX equation
    :return: Word equation
    """
    import os, latex2mathml.converter
    from lxml import etree
    mathml = latex2mathml.converter.convert(latex_input)
    tree = etree.fromstring(mathml)
    script_path = os.path.dirname(os.path.abspath(__file__))
    xslt = etree.parse(os.path.join(script_path,'..','Shared','MML2OMML.XSL'))
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    return new_dom.getroot()

def payback(ACS, IC) -> str:
    """
    Format payback period by year and month
    :param ACS: Annual Cost Savings ($/yr)
    :param IC: Implementation Cost ($)
    :return: formatted Payback Period (str)
    """
    import math
    PB = IC / ACS
    if PB <= 11.0 / 12.0:
        PB = math.ceil(PB * 12.0)
        PBstr = str(PB) + " month"
    else:
        PB = math.ceil(PB * 10.0) / 10.0
        PBstr = str(PB).rstrip("0").rstrip(".") + " year"
    if PB > 1.0:
        PBstr = PBstr + "s"
    return PBstr