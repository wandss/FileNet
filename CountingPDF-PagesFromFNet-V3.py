# -*- coding: cp1252 -*-
import sqlite3,os,time,getpass,pyPdf


class Begin(object):
    """
    This Class creates a database using SQLITE3 to store configuration
    for connecting into CMIS repository.

    If there is a Data Base with more than one configuration,
    the user will be prompted to choose which configuration is the correct.
    """
    def __init__(self):
        dbs = os.listdir(os.getcwd())
        self.db = ''
        self.opt = 0
        for d in dbs:
            if d =='confs.db':
                self.db = d
        if self.db == '':
            print('No configuration or Data Base Found:')
            self.createDb()
        else:
            rs = DB.listAll()
            if len(rs)>1:
                print ("More than one Configuration found. Choose one to be used:")
                for i,r in enumerate(rs):
                    print(str(i+1)+' > '+str(r[1])+' | '+str(r[2]))
                self.opt = int(raw_input('\n> '))                               
            
            
    def createDb(self):
        """
        This method creates a new Data Base.
        """
        self.db = DB.connect()
        self.cursor = self.db.cursor()
        self.cursor.execute("CREATE TABLE configuration(\
                        id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,\
                        url TEXT NOT NULL,\
                        user TEXT NOT NULL,\
                        prop TEXT NOT NULL,\
                        create_date DATE NOT NULL);")
        print('Creating table.....')
        self.db.close()
        self.populateDb()
    def populateDb(self):
        """
        This method the new Data Base.
        """
        self.db = DB.connect()
        self.cursor = self.db.cursor()
        self.url = raw_input("Inform the URL for the CMIS Compliant Repository:\n>")
        self.user  = raw_input("Inform the USER to connect to the repository:\n>")
        self.prop = raw_input("Inform the Property Template to store the number of pages:\n>")
        cd = time.localtime()
        self.create_date = str(cd.tm_year)+'-'+str(cd.tm_mon)+'-'+str(cd.tm_mday)
        self.cursor.execute("INSERT INTO configuration(url,user,prop,create_date)\
                            VALUES('%s','%s','%s','%s');"%(self.url,
                                                           self.user,
                                                           self.prop,
                                                           self.create_date))
        self.db.commit()
        self.db.close()
        print("Data succesfully saved!")       

class DB(object):
    @staticmethod
    def connect():
        """
        This method belongs to DB class. It connects
        to the Data Base and returns a sqlite3.Connection object.
        """
        try:
            db = sqlite3.connect('confs.db')
            return db
        except Exception as e:            
            print str(e)
            print ("Error while creating a new Data Base!!!\
                   \nCertify that you have permissions to write in\
                   \ndirectory: %s"%str(os.getcwd()))
            exit()
    @staticmethod
    def listAll():
        """
        This method returns a result set
        with all the entries in the Data Base.
        """
        rs = []
        cursor = DB.connect().cursor()
        results = cursor.execute('select * from configuration')
        for r in results:
            rs.append(r)
        return (rs)
        

class Conecta(object):
    """
    This class returns a Cmis.Repository object.
    """
    @staticmethod	
    def cmisConnect(opt=0):
        """
        This method connects to a CMIS Compliant repository using
        configuration from the Data Base.
        """
        try:                
            from cmislib.model import CmisClient
            rs = DB.listAll()            
            url = rs[opt-1][1].encode('utf8')
            user = rs[opt-1][2].encode('utf8')
            passwd = getpass.getpass('Password for USER: "%s" to connect to the Repository:\n>'%user)
            client = CmisClient(url,user,passwd)
            rep = client.defaultRepository
            return (rep)

        except Exception as e:  
            print ('Exception>>> '+str(e)) 
            message = "Could'nt connect to Repository:\nPlease check user and password"
            print (message)
            return (message)

class StartCounting(object):
    """
    This is the main class.
    Here the user will be prompted to inform the ID's from the documents in
    the repository. The ID's can be retrieved from the Repository Data Bases,
    or from user interface like IBM WorkplaceXT or IBM Content Navigator.

    The user will also choose if the number of pages from the PDF
    files will be updated in the repository.
    """

    def __init__(self):
        self.program = Begin()
        self.rep = Conecta.cmisConnect(self.program.opt) 
        self.pages = 0       
        self.errors = []
        print("Update number of pages in FileNet:\n1: YES")
        self.opt = raw_input('> ')  

        while True:    
            print ("Iform ID(s):\n'0' for EXIT:")
            self.texto = raw_input('> ')
            if self.texto in'0':
                break
            else:
                self.Ids = self.texto.split('\n')    
                if self.Ids[0][9:10]=='-' or self.Ids[0][8:9]=='-':
                    self.ids = self.fnetIds(self.Ids)
                else:
                    self.ids = self.dbIds(self.Ids)        
                for i in self.ids:
                    print (i)
                    self.docs = self.getDocs(i,self.rep)
                    if str(type(self.docs)).startswith('<class'):
                        self.errors.append(self.writeDocs(self.docs))
                        self.pages += self.countPages(self.docs,self.opt,self.program.opt)
                    else:   
                        print (self.docs)
        self.printErrors(self.errors)
        
    def dbIds(self,Ids):
        """
        This method creates a list with document IDs when
        they were retrieved from the Repository's Data Base.        
        """
        ids = []
        for i in Ids:                
            ids.append('idd_'+i[6:8]+i[4:6]+i[2:4]+i[0:2]+'-'+i[10:12]+i[8:10]+'-'+i[14:16]+i[12:14]+'-'+i[16:20]+'-'+i[20:32])        
        return ids
    
    def fnetIds(self,Ids):
        """
        This method creates a list with document IDs when
        they were retrieved from user interfaces like:
        IBM's WorkplaceXT or IBM's Content Navigator.
        """
        ids = []        
        for i in Ids:            
            if i.startswith('{') or i.endswith('}'):            
                if i.startswith('{'):
                    ids.append('idd_'+i[1:])
                elif i.endswith('}'):
                    ids.append('idd_'+i[:-1])
                elif i.startswith('{') and i.endswith('}'): 
                    ids.append('idd_'+i[1:-1])
            else:
                ids.append('idd_'+i)        
        return ids
     
    def getDocs(self,ids,rep):
        """
        This method gets the documents from the repository,
        using the ids passed from the users.
        """
        try:
            docs = rep.getObject(ids)
            return docs
        except Exception as e:                
            docs = ('Error: '+str(e))
            return docs        
            print(str(e))
     
    def writeDocs(self,docs):
        """
        This method downloads the documents from the repository
        to a local directory so the amount of pages can be extracted
        from it.
        """
        try:
            f = open(docs.name,'wb')
            f.write(docs.getContentStream().read())
            f.close()
        except Exception as e:        
            print ('>>>> '+str(e))
            print(docs.name)
            #os.remove(f.name)
            return str(e)+' - '+docs.name
     
    def countPages(self,doc,opt,dbopt):
        """
        This method extracts the number of pages from a
        PDF file, downloaded from the repository and then
        excludes this file from the local directory.

        If the user has chosen to update the number of pages in
        the repository, this method will do so.
        """
        f = open(doc.name,'rb')
        try:
            pdf = pyPdf.PdfFileReader(f)
            pages = pdf.getNumPages()
            f.close()
            os.remove(f.name)
            print('Total Number of Pages From "%s": %d'%(doc.name,pages))
            if opt ==str(1):                
                rs = DB.listAll()                
                doc.updateProperties({rs[dbopt][3]:pages})
            return pages
        except Exception as e:        
            print ('\nNone PDF file Returned:\n'+str(e))
            return (0)
     
    def printErrors(self,lista):
        for l in lista:
            if l!=None:
                print(l)
                print self.pages
        print ("Total amount of pages from all the files: %d"%self.pages)
    
def info():
    print(37*'* ')
    print ("""*                                                                       *
*       The IBM's FileNet P8 System, doesn't automatically stores the   *
*   total number of pages from a PDF file.                              *
*                                                                       *
*       This programs intends to solve this by couting the number of    *
*   pages from a PDF that is already stored in the IBM's FileNet P8,    *
*   and save (if wished so) this information into a 'Property Template'.*
*                                                                       *
*       In order to run this program it will be necessary to have:      *
*           - IBM FileNetP8                                             *
*           - IBM CMIS                                                  *
*           - Apache Open CMIS (cmislib)                                *
*           - PyPDF                                                     *
*       Open CMIS and PyPDF can be installed with pip.                  *
*                                                                       *
*       In order to have the total number of pages saved in FileNet     *
*   (which might be a really important information), you'll have,       *
*   firstly, to create a 'Property Template' in IBM's FileNet and       *
*   assotiate this property template to the whished Document classes.   *
*                                                                       *
*       In this program, you'll have to inform the url address to       *
*   connect to the Repository, usually this addres is something like:   *
*                                                                       *
*        http://servername:port/cmisservice/resources/Service           *
*                                                                       *         
*       For more information about this, verfy both Apache CMIS Lib     *
*   and IBM CMIS documentation.
*
*       For the first time running this program it will create a
*   SQLLite3 Data Base to store:

        - URL addres for CMIS,
        - Username from FileNet (recomend to use FileNet
            P8 Domain Admin)
        - Symbolic Name from the Property Template

*       The program will prompt the user to inform the ID from
    documents stored in FileNet repository.
    These ID's can be obtainned from FileNet's database or directly
    from the document properties

    +++++ falar sobre os tipos de IDs
    
""")
    print(37*'* ')
    
        
#info()
StartCounting()     
"""
TODO:
 - Alterar para que diferentes tipos de ID {8A9434C3-63E2-4FC1-B031-67CB65029464} ou
    ou C7D649A6A71D3846A438DF5233D419D2
    sejam carregados em uma mesma busca
 - Testar erros de conexão ao repositório
 - Criar menu de instruções
"""





    

