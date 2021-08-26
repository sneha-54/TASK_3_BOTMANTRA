#!/usr/bin/env python
# coding: utf-8

# In[1]:


from docx import Document
from lxml import etree
import zipfile
ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
#Function to extract all the comments of document(Same as accepted answer)
#Returns a dictionary with comment id as key and comment string as value
def get_document_comments(docxFileName):
    comments_dict={}
    docxZip = zipfile.ZipFile(docxFileName)
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment',namespaces=ooXMLns)
    for c in comments:
        comment=c.xpath('string(.)',namespaces=ooXMLns)
        comment_id=c.xpath('@w:id',namespaces=ooXMLns)[0]
        comments_dict[comment_id]=comment
    return comments_dict
#Function to fetch all the comments in a paragraph
def paragraph_comments(paragraph,comments_dict):
    comments=[]
    for run in paragraph.runs:
        comment_reference=run._r.xpath("./w:commentReference")
        if comment_reference:
            comment_id=comment_reference[0].xpath('@w:id',namespaces=ooXMLns)[0]
            comment=comments_dict[comment_id]
            comments.append(comment)
    return comments
#Function to fetch all comments with their referenced paragraph
#This will return list like this [{'Paragraph text': [comment 1,comment 2]}]
def comments_with_reference_paragraph(docxFileName):
    document = Document(docxFileName)
    comments_dict=get_document_comments(docxFileName)
    comments_with_their_reference_paragraph=[]
    for paragraph in document.paragraphs:  
        if comments_dict: 
            comments=paragraph_comments(paragraph,comments_dict)  
            if comments:
                comments_with_their_reference_paragraph.append({paragraph.text: comments})
    return comments_with_their_reference_paragraph
if __name__=="__main__":
    document="highlights.docx"  #filepath for the input document
    com=comments_with_reference_paragraph(document)
    l=[]
    for i in com:
        l.extend(list(i.values())[0])
    print('Comments are:',l,'\nNo. of comments are:',len(l))
    char_count=0  #character count in comment with spaces
    space_count=0  #it will calculate word count
    for i in l:
        space_count+=1
        char_count+=len(i)
        for j in i:
            if j==" ":
                space_count+=1
    print('No. of characters with spaces',char_count) 
    print('No. of words',space_count) 
    from textblob import TextBlob
    text=" ".join(l)
    lang = TextBlob(text)
    l=lang.detect_language()
    from pycountry import languages
    lang_name = languages.get(alpha_2=l).name
    print(lang_name)
    


# In[ ]:




