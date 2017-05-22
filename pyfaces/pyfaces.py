import  sys
from string import split
from os.path import basename
import eigenfaces
import platform
import os


class PyFaces:
    def __init__(self,testimg,imgsdirFeliz,imgsdirNeutro,imgsdirMolesto,egfnum,thrsh):
        self.testimg=testimg
        self.matchfile = None
        self.arregloEmociones = []
        self.arregloDistancias = []
        self.imgsdirFeliz=os.path.abspath(imgsdirFeliz)
        self.imgsdirNeutro=os.path.abspath(imgsdirNeutro)
        self.imgsdirMolesto=os.path.abspath(imgsdirMolesto)
        self.threshold=thrsh
        self.egfnum=egfnum        
        parts = split(basename(testimg[1]),'.')
        extn=parts[len(parts) - 1]   
        self.facet= eigenfaces.FaceRec()
        self.egfnum=self.set_selected_eigenfaces_count(self.egfnum,extn)
        self.facet.checkCache(self.imgsdirFeliz,extn,self.imgnamelist,self.egfnum,self.threshold)
        for j in range(len(self.testimg)):
            self.mindist,self.matchfile=self.facet.findmatchingimage(self.testimg[j],self.egfnum,self.threshold)
            if self.mindist < 1e-10:
                self.mindist=0
            self.emocionMenor = self.mindist
            if platform.system()=='Windows':
                self.emocionFinal = split(self.matchfile,'\\')
            else:
                self.emocionFinal = split(self.matchfile,'/')
            self.emocionFinal = self.emocionFinal[len(self.emocionFinal)-2]
            self.arregloEmociones.append(self.emocionFinal)
            self.arregloDistancias.append(self.mindist)
            
    def set_selected_eigenfaces_count(self,selected_eigenfaces_count,ext):               
        self.imgnamelist=self.facet.parsefolder(self.imgsdirFeliz,ext)
        self.imgnamelist2=self.facet.parsefolder(self.imgsdirNeutro,ext)
        self.imgnamelist3=self.facet.parsefolder(self.imgsdirMolesto,ext)
        self.imgnamelist=self.imgnamelist+self.imgnamelist2+self.imgnamelist3
        numimgs=len(self.imgnamelist)
        if(selected_eigenfaces_count >= numimgs  or selected_eigenfaces_count == 0):
            selected_eigenfaces_count=numimgs/2
        else:
            selected_eigenfaces_count=selected_eigenfaces_count
        print selected_eigenfaces_count
        return selected_eigenfaces_count
        
        
