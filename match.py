class Match:
    def __init__(self):
        self.complete_match = []
        self.partially_match = []
        self.mismatch = []
        self.complete_match_dct = {}
        self.partially_match_dct = {}
        self.mismatch_docx = {}
        self.mismatch_xlsx = {}
        # xls_data is a dictionary, values are lists
        # docx_data is a list which consists several dictionaries


    def match(self, docx_data, xls_data):
        print("data from excel file: ", xls_data)
        print("data from docx file: ", docx_data)
        self.xls_data_mismatch = copy.deepcopy(xls_data)
        self.docx_data_mismatch = copy.deepcopy(docx_data)

        # first I get a first element of docx_data and compare it with xlsx_data
        for docx_semester, docx_subjects in enumerate(docx_data):
            #print(docx_semester, docx_subjects)
            # checking every subject in docx_data if it is in xls_data
            for docx_subject, docx_info in docx_subjects.items():
                #print(docx_subject)
                if docx_subject in xls_data:
                    xls_subject = docx_subject
                    self.xls_data_mismatch.pop(xls_subject, None)
                    self.docx_data_mismatch[docx_semester].pop(docx_subject, None)
                else:
                    xls_subject = None
                #print(docx_subject, xls_subject)




                if xls_subject:

                    # information for dictionary, if subjects are matching
                    hours = xls_data[xls_subject][0]
                    kr_type_xls = xls_data[xls_subject][1]
                    kr_type_docx = docx_data[docx_semester][docx_subject][0]
                    # you should add information about semester from excel file
                    mark = docx_data[docx_semester][docx_subject][1]
                    date = docx_data[docx_semester][docx_subject][2]


                    if docx_subject == xls_subject and docx_info[0] == xls_data[xls_subject][1]:

                        # checking if the subject is already in the list or not
                        if (docx_subject, xls_subject) in self.complete_match:
                            # find the index of the subject
                            index = self.complete_match.index((docx_subject, xls_subject))

                            # replace the subject
                            self.complete_match[index] = (docx_subject, xls_subject)

                            # appending elements to the dictionary (maybe in future i will change this part of code)
                            self.complete_match_dct[docx_subject] = (hours, kr_type_xls, mark, date)

                        else:
                            self.complete_match.append((docx_subject, xls_subject))

                            # appending elements to the dictionary
                            self.complete_match_dct[docx_subject] = (hours, kr_type_xls, mark, date)
                    else:
                        if (docx_subject, xls_subject) in self.partially_match:
                            # find the index of the subject
                            index = self.partially_match.index((docx_subject, xls_subject))

                            # replace the subject
                            self.partially_match[index] = (docx_subject, xls_subject)

                            # appending to the final dictionary all the information about subjects
                            self.partially_match_dct[docx_subject] = (hours, kr_type_xls, kr_type_docx, mark, date)
                        else:
                            self.partially_match.append((docx_subject, xls_subject))

                            # appending to the final dictionary all the information about subjects
                            self.partially_match_dct[docx_subject] = (hours, kr_type_xls, kr_type_docx, mark, date)
                else:
                    self.mismatch.append(docx_subject)
        #return self.complete_match, self.partially_match, self.mismatch
        return self.complete_match_dct, self.partially_match_dct, self.xls_data_mismatch, self.docx_data_mismatch
