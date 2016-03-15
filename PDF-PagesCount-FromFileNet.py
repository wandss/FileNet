# -*- coding: cp1252 -*-
import sqlite3,os,time,getpass


class Menu(object):

    @staticmethod
    def show():
        
        
        while True:            
            print('Available Options:\n')
            print('1 > Count and Update Pages on Repository.')
            print('2 > Just count Pages, whithout Update.')
            print('3 > Add configuration into Data Base.')
            print('4 > List configuration stored in Data Base.')
            print('5 > See documentation.')
            print('0 > Exit')
            opt = raw_input('\n\n>> ')
            if opt=='0':
                print('Exiting....')
                break
            if opt=='1' or opt=='2':
                StartCounting(opt)
            if opt=='3':
                Begin(opt).populateDb(opt)
                #db.populateDb()
            if opt=='4':
                rs = DB.listAll()
                if len(rs)==0:
                    print('\nThere is no Data in Database')
                    opt = 3
                    Begin(opt).populateDb(opt)                    
                print('\n')
                for r in rs:
                    print(str(r[0])+' | '+str(r[1])+' | '+str(r[2])+' | '+str(r[3]))
                print('\n')
            if opt=='5':
                info()
                print('\n')            
            else:
                print opt
                print('Invalid option!!!! \n')
                

        exit()
                
            

class Begin(object):
    """
    This Class creates a database using SQLITE3 to store configuration
    for connecting into CMIS repository.

    If there is a Data Base with more than one configuration,
    the user will be prompted to choose which configuration is the correct.
    """
    def __init__(self,uopt):        
        dbs = os.listdir(os.getcwd())
        self.db = ''
        self.opt = 0
        self.useropt = uopt
        for d in dbs:
            if d =='confs.db':
                self.db = d
        if self.db == '':
            print('No configuration or Data Base Found:')
            self.createDb()
        else:
            rs = DB.listAll()
            if len(rs)>1 and self.useropt!='3':
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
        self.populateDb(self.opt)
    def populateDb(self,opt):
        """
        This method the new Data Base.
        """
        print opt
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
        print opt
        if opt == 0:
            StartCounting(opt)
        else:
            Menu.show()

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
            MenuShow()
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

class StartCounting(object):
    """
    This is the main class.
    Here the user will be prompted to inform the ID's from the documents in
    the repository. The ID's can be retrieved from the Repository Data Bases,
    or from user interface like IBM WorkplaceXT or IBM Content Navigator.

    The user will also choose if the number of pages from the PDF
    files will be updated in the repository.
    """
    

    def __init__(self,opt):        
        self.program = Begin(opt)
        self.rep = Conecta.cmisConnect(self.program.opt) 
        self.pages = 0       
        self.errors = []             
        self.opt = opt
        print self.opt

        while True:    
            print ("Inform ID(s):\n'0' for RETURN TO MAIN MENU:")
            self.texto = raw_input('> ')
            if self.texto in'0':
                self.printResults(self.errors)
                Menu.show()
            else:
                self.Ids = self.texto.split('\n')    
                self.ids = self.formatIds(self.Ids)     
                for i in self.ids:
                    print (i)
                    self.docs = self.getDocs(i,self.rep)
                    if str(type(self.docs)).startswith('<class'):
                        self.errors.append(self.writeDocs(self.docs))
                        self.pages += self.countPages(self.docs,self.opt,self.program.opt)
                    else:   
                        print (self.docs)        
        
    def formatIds(self,Ids):
        """
        This method creates a list with document IDs when
        they were retrieved from the Repository's Data Base.        
        """
        ids = []
        for i in Ids:
            if i[9:10]=='-' or i[8:9]=='-':
                if i.startswith('{') or i.endswith('}'):            
                    if i.startswith('{'):
                        ids.append('idd_'+i[1:])
                    elif i.endswith('}'):
                        ids.append('idd_'+i[:-1])
                    elif i.startswith('{') and i.endswith('}'): 
                        ids.append('idd_'+i[1:-1])
                else:
                    ids.append('idd_'+i)
            else:
                ids.append('idd_'+i[6:8]+i[4:6]+i[2:4]+i[0:2]+'-'+i[10:12]+i[8:10]+'-'+i[14:16]+i[12:14]+'-'+i[16:20]+'-'+i[20:32])        
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
        import pyPdf
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
     
    def printResults(self,lista):
        for l in lista:
            if l!=None:
                print(l)
                print self.pages
        print ("Total amount of pages from all the files: %d"%self.pages)

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
            message = ("Could'nt connect to Repository:\
                        \nPlease check user and password!\
                        \nAlso verify the DataBase information!")
            print (message)
            Menu.show()
    
def info():
    print(37*'* ')
    print ("""*                                                                       *
*   This Program was developed by Wanderley de Souza.                   *
*                   wandss@gmail.com                                    *                *
*                                                                       *
*   Like every program, this code can always be improved. If you        *
*   find any bugs, feel free to contact me as well if you have any      *
*   doubts.                                                             *
*                                                                       *
*   Introduction:                                                       *
*                                                                       *
*       As far as I concern, the IBM's FileNet P8 System, doesn't       *
*   automatically stores the total number of pages from a PDF file,     *
*   while saving files in the Repository.                               *                                        *
*       This programs intends to solve this by counting the number of   *
*   pages from a PDF that is already stored in the IBM's FileNet P8,    *
*   and save (if wished to do so) this information into a 'Property     *
*   Template'.                                                          *
*                                                                       *
*   Requirements:                                                       *
*                                                                       *
*       In order for this program to run, it will be necessary          *
*    to have:                                                           *
*           - IBM FileNetP8                                             *
*           - IBM CMIS                                                  *
*           - Apache Open CMIS (cmislib)                                *
*           - PyPDF                                                     *
*       Apache Open CMIS and PyPDF can be installed with pip.           *
*                                                                       *
*       In order to have the total number of pages saved in FileNet     *
*   (which might be a really usefull information) you'll have,          *
*   firstly, to create a 'Property Template' in IBM's FileNet and       *
*   assotiate this property template to the whished Document classes.   *
*                                                                       *   
*   How it works:                                                       *
*                                                                       *
*       In this program, you'll have to inform the URL address to       *
*   connect to the Repository. Usually this addres is something like:   *
*                                                                       *
*        http://servername:port/cmisservice/resources/Service           *
*                                                                       *         
*       For more information about this, verfy both Apache CMIS Lib     *
*   and IBM CMIS documentation.                                         *
*                                                                       *
*       For the first time running this program it will create a        *
*   SQLLite3 Data Base to store:                                        *
*                                                                       *
*       - URL addres for CMIS compliant Repository.                     *
*       - USERNAME for connecting in FileNet.                           *
*           (recomend to use FileNet P8 Domain Admin)                   *
*       - Property Template's Symbolic Name.                            *
*                                                                       *
*       The program will prompt the user to inform the ID from the      *
*   documents stored in FileNet repository. These ID's can be           *
*   obtainned from FileNet's database or directly from the document's   * 
*   properties.                                                         *
*                                                                       *
*   Data Base ID: E541220C9058F2489FF2AEF63D53DE61                      *
*   Dcument's properties ID: {5B8E204C-B822-4793-964F-1F83EDAD830A}     *
*                                                                       *
*       You can pass in the program, one ID at a time or as many        *
*   IDs you wish, like:                                                 *
*                                                                       *
*       One ID: E541220C9058F2489FF2AEF63D53DE61                        *
*                                                                       *        
*       Many IDs:52BB10B65FD47648834FF61F9F765BA6                       *
*                1E908F5A1F19044B811521E78BE935C7                       *
*                F8B9C5AE3245D5448781C26836234F4C                       *
*                D1343C1CD525E74E8CA8B70675F80292                       *
*                E66BD9596A9D5E48B6BD7DA597A444D1                       *
*                E6E253FFFB60C040B96CB87AD65D4943                       *
*                F8798BA696BE09459B9986530A540EAF                       *
*                                                                       *
*       It is also possible to pass a list of IDs with mixed type       *
*  of IDs, like:                                                        *
*                                                                       *
*               1E908F5A1F19044B811521E78BE935C7                        *
*               F8B9C5AE3245D5448781C26836234F4C                        *
*               {5B8E204C-B822-4793-964F-1F83EDAD830A}                  *
*               {93EB4EF2-192F-4439-A8AF-8BAD40F40400}                  *
*               F7B2C6819056DA42891A24980C410C7C                        *
*               FF5F85D8FD68DA4AB36F10E6BEC3C408                        *
*               {0B22EB4C-E83C-49E2-8279-2209FF6E684C}                  *
*               DC7BA46418429348941894C772993886                        *
*               E4ADADA16C2D3E48AA5A629DB394965E                        *
*                                                                       *
*       Running the Program:                                            *
*                                                                       *
*           The program can be started from a prompt or a shell         * 
*       terminal. I recommend though, that you run this program         *
*       from IDLE, since in IDLE, you can paste a list of ID's as       *
*       shown above.                                                    *
*                                                                       *
*            While running, the program will print the name of the      *
*        document and it's amount of pages.                             *
*            At the end of the counting, the program will inform the    *
*        total number of pages for all the passed Documents.            * 
*                                                                       *
*    Final considerations:                                              *
*                                                                       *
*        As mentioned before, feel free for making questions,           *
*    suggestions and for contacting me.                                 *
*        Your considerations are very important for the improvement     *
*    of this tool and for my own improvement as well.                   * 
*                                                                       *
*    Thanks for the attetion and I really hope that this program        *
*    helps you as much it is helping me.                                *
*                                                                       *
*    Regards,                                                           *
*                                                                       *
*    Wanderley de Souza                                                 *
""")
    print(37*'* ')
    
if __name__=='__main__':
    Menu.show()
    
"""
TODO:
 - Criar menu de instruções
"""





    

