"""
(Purpose) IAC.py is a module that contains functions used in the IAC report
"""

import math, os, locale
from lxml import etree
import latex2mathml.converter

def combine_words(words):
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

def grouping_num(dic):
    """
    Add thousand separator to numbers in a dictionary and format it to string
    :param dic: Dictionary
    :return: Dictionary with keys in thousand separator
    """
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

def add_eqn(doc, eqn, eqn_input):
    """
    Add equation to Word document, search for eqn in doc and replace with eqn_input
    :param doc: Document
    :param eqn: Equation string
    :param eqn_input: Equation input
    :return: None
    """
    for p in doc.paragraphs:
        if eqn in p.text:
            p.text = p.text.replace(eqn, '')
            word_math = latex2word(eqn_input)
            p._element.append(word_math)

def latex2word(latex_input):
    """
    Convert LaTeX equation to Word equation
    :param latex_input: LaTeX equation
    :return: Word equation
    """
    mathml = latex2mathml.converter.convert(latex_input)
    tree = etree.fromstring(mathml)
    script_path = os.path.dirname(os.path.abspath(__file__))
    xslt = etree.parse(os.path.join(script_path,'..','Shared','MML2OMML.XSL'))
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    return new_dom.getroot()

def payback(ACS,IC):
    """
    Formet payback period by year and month
    :param ACS: Annual Cost Savings ($/yr)
    :param IC: Implementation Cost ($)
    :return: formatted Payback Period (str)
    """
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