import numpy as np
import math
import sph
from srtm import get_elevation


R = sph.a_e

def DMS_in_DD(deg,mn,sec):
    """ Градусы/минуты/секкунды => градусы """
    d=deg+(mn/60)+sec/3600
    return d

def DD_in_DMS(d):
    """ Градусы => градусы/минуты/секкунды """
    deg = int(d)
    mn = int((d-deg)*60)
    sec = ((d-deg)*60-mn)*60
    return (deg,mn,round(sec,2))


def Azimyt1(point1,point2):
    fi1 = point1[0]/180*np.pi
    lm1 = point1[1]/180*np.pi
    fi2 = point2[0]/180*np.pi
    lm2 = point2[1]/180*np.pi
    A = np.arctan2(np.sin(lm2-lm1)*np.cos(fi2),np.cos(fi1)*np.sin(fi2)-np.cos(fi2)*np.sin(fi1)*np.cos(lm2-lm1))/np.pi*180
    return A

def Azimyt2(point1,point2):
    fi1 = point1[0]/180*np.pi
    lm1 = point1[1]/180*np.pi
    fi2 = point2[0]/180*np.pi
    lm2 = point2[1]/180*np.pi
    a = 6378.2  
    b = 6356.9
    e = np.sqrt(1-b**2/a**2)
    n2 = np.log(np.tan(np.pi/4+fi2/2)*((1-e*np.sin(fi2))/(1+e*np.sin(fi2)))**(e/2))
    n1 = np.log(np.tan(np.pi/4+fi1/2)*((1-e*np.sin(fi1))/(1+e*np.sin(fi1)))**(e/2))
    A = np.arctan((lm2-lm1)/(n2-n1))/np.pi*180
    return A

def greatCircleDistance(kord1, kord2,R_eath=R):
    """ Расчёт растояния по географическим координатам """
    (lat1,lon1) = kord1
    (lat2,lon2) = kord2
    def haversin(x):
        return np.sin((x/180*np.pi)/2)**2 
    return R_eath * (2 * np.arcsin(np.sqrt(haversin(lat2-lat1) 
          +np.cos(lat1/180*np.pi) * np.cos(lat2/180*np.pi) * haversin(lon2-lon1))))


def OnePoint(point, azi, dist):
    lat1 = math.radians(point[0])
    lon1 = math.radians(point[1])
    azi = math.radians(azi)
    dist = dist / sph.a_e
    lat2, lon2 = sph.direct(lat1, lon1, dist, azi)
    #dist, azi2 = sph.inverse(lat2, lon2, lat1, lon1)
    #print ("%f %f %f" % (math.degrees(lat2), math.degrees(lon2), math.degrees(azi2)))
    return (math.degrees(lat2),math.degrees(lon2))
    
def WritePoint(point1,point2,step):
    list_step_points = []
    list_step_elevations = []

    dist = greatCircleDistance(point1,point2)
    azi = Azimyt1(point1,point2)
    l = int(dist/step)
    list_dist = [(i+1)*step  for i in range(l) if i!=l-1]

    list_step_points.append(point1)
    el = get_elevation(point1[0],point1[1])
    if el==-32768: raise Exception("Elevation not found")
    list_step_elevations.append(el)

    for d in list_dist:
        p = OnePoint(point1, azi, d)
        list_step_points.append(p)
        el = get_elevation(p[0],p[1])
        if el==-32768: raise Exception("Elevation not found")
        list_step_elevations.append(el)
    
    list_step_points.append(point2)
    el = get_elevation(point2[0],point2[1])
    if el==-32768: raise Exception("Elevation not found")
    list_step_elevations.append(el)
    
    return np.array(list_dist,dtype=np.float64), np.array(list_step_elevations,dtype=np.float64),\
         dist, np.array(list_step_points,dtype=np.float64)


def Grafik(point1,point2,step=0.030):
    """ Строим кривые рельефа для графика """
    list_dist, list_step_elevations, dist, list_step_points = WritePoint(point1,point2,step)
    l = len(list_step_elevations)
    alf = dist/R
    Xrd = 2*R*np.sin(alf/2)

    bt = list_dist*alf/dist-alf/2

    lx_ev1 = np.array([-Xrd/2,(-(list_step_elevations[0])/1000*np.sin(alf/2)-Xrd/2)],dtype=np.float64)
    ly_ev1 = np.array([0,(list_step_elevations[0])*np.cos(alf/2)],dtype=np.float64)

    lx_ev2 = np.array([Xrd/2,((list_step_elevations[l-1])/1000*np.sin(alf/2)+Xrd/2)],dtype=np.float64)
    ly_ev2 = np.array([0,(list_step_elevations[l-1])*np.cos(alf/2)],dtype=np.float64)

    lx_R = R*np.sin(bt)
    ly_R = R*1000*np.cos(bt)-R*1000*np.cos(alf/2)
    lx_ex = (R+list_step_elevations[1:l-1]/1000)*np.sin(bt)
    ly_ex = (R*1000+list_step_elevations[1:l-1])*np.cos(bt)-R*1000*np.cos(alf/2)

    def hh1(h):
        return np.array([(-(h+list_step_elevations[0])/1000*np.sin(alf/2)-Xrd/2),\
            (h+list_step_elevations[0])*np.cos(alf/2)],dtype=np.float64) 
    def hh2(h):
        return np.array([((h+list_step_elevations[l-1])/1000*np.sin(alf/2)+Xrd/2),\
            (h+list_step_elevations[l-1])*np.cos(alf/2)],dtype=np.float64)  
    def h_min1(y):
        return y/np.cos(alf/2)-list_step_elevations[0]
    def h_min2(y):
        return y/np.cos(alf/2)-list_step_elevations[l-1]
    
    f1 = lambda h:hh1(h)
    f2 = lambda h:hh2(h)
    f3 = lambda y:h_min1(y)
    f4 = lambda y:h_min2(y)

    A1 = ly_ev1[0] - ly_ev1[1]
    B1 = lx_ev1[0] - lx_ev1[1]
    C1 = lx_ev1[0]*ly_ev1[1]-lx_ev1[1]*ly_ev1[0]

    A2 = ly_ev2[0] - ly_ev2[1]
    B2 = lx_ev2[0] - lx_ev2[1]
    C2 = lx_ev2[0]*ly_ev2[1]-lx_ev2[1]*ly_ev2[0]

    return (Xrd, dist,\
        (lx_ev1,ly_ev1,lx_ev2,ly_ev2,lx_R,ly_R,lx_ex,ly_ex), (f1,f2,f3,f4),\
        (A1,B1,C1), (A2,B2,C2), (list_step_elevations[0],list_step_elevations[l-1]),\
            list_step_points,list_dist,list_step_elevations)

def h_min(lx_ev,ly_ev,xy,Yr,zapas,f=lambda y:y):
    """ Определение минимальной высоты вышки относительно другой """
    A = xy[1]-(ly_ev+zapas)
    B = xy[0]-lx_ev
    C = xy[0]*(ly_ev+zapas)-lx_ev*xy[1]

    y = np.max((Yr[0]*C-A*Yr[2])/(Yr[0]*B-A*Yr[1]))
    return f(y)

def Frenel(lx_ev,xy1,xy2,f,ps):
    
    alf = np.arctan((xy2[1]-xy1[1])/(xy2[0]-xy1[0])/1000)
    dist = (xy2[0]-xy1[0])*np.cos(alf)
    S = (lx_ev/1000-xy1[0])*np.cos(alf)
    #f = 11 # ГГц
    Rg = 17.3 * np.sqrt(S*(dist-S)/(f*dist))*(ps/100) / np.cos(alf)
    ly_ev = -(xy2[0]*1000-lx_ev) * np.tan(alf) + xy2[1]

    #1
    lx_f1 = (lx_ev/np.tan(np.pi/2-alf)+lx_ev/np.tan(alf)+ly_ev-(ly_ev-Rg))/(1/np.tan(alf)+1/np.tan(np.pi/2-alf))
    ly_f1 = (ly_ev/np.tan(np.pi/2-alf)+(ly_ev-Rg)/np.tan(alf)+lx_ev-lx_ev)/(1/np.tan(alf)+1/np.tan(np.pi/2-alf))

    lx_f2 = (lx_ev/np.tan(np.pi/2-alf)+lx_ev/np.tan(alf)-(ly_ev+Rg)+ly_ev)/(1/np.tan(alf)+1/np.tan(np.pi/2-alf))
    ly_f2 = (ly_ev/np.tan(np.pi/2-alf)+(ly_ev+Rg)/np.tan(alf)+lx_ev-lx_ev)/(1/np.tan(alf)+1/np.tan(np.pi/2-alf))


    return lx_f1/1000,ly_f1,lx_f2/1000,ly_f2

def H_Min(lx_ev,ly_ev,xy,Yr,fg,ps,h,f=lambda y:y):
    if h == "h2":
        dist = np.abs(xy[0]*2)
        S = (lx_ev/1000+dist/2)

        Rm = 17.3 * np.sqrt(S*(dist-S)/(fg*dist))*(ps/100) *1.1

        gama = np.pi/2-np.arctan((Rm)/(lx_ev-xy[0]*1000))
        beta  = np.arctan((ly_ev-xy[1])/(lx_ev-xy[0]*1000))

        fi=gama-beta-np.pi/2

        lx = (lx_ev+Rm*np.sin(fi))/1000
        ly = ly_ev+Rm*np.cos(fi)

        A = xy[1]-ly
        B = xy[0]-lx
        C = xy[0]*(ly)-lx*xy[1]

        y = np.max((Yr[0]*C-A*Yr[2])/(Yr[0]*B-A*Yr[1]))

        return f(y)
    elif h=="h1":
        dist = np.abs(xy[0]*2)
        S = (dist/2-lx_ev/1000)
        
        Rm = 17.3 * np.sqrt(S*(dist-S)/(fg*dist)) *1.1
        gama = np.pi/2-np.arctan((Rm)/(lx_ev-xy[0]*1000))

        beta  = np.arctan((ly_ev-xy[1])/(lx_ev-xy[0]*1000))

        fi=gama-beta-np.pi/2
        lx = (lx_ev+Rm*np.sin(fi))/1000
        ly = ly_ev+Rm*np.cos(fi)

        A = xy[1]-ly
        B = xy[0]-lx
        C = xy[0]*(ly)-lx*xy[1]

        y = np.max((Yr[0]*C-A*Yr[2])/(Yr[0]*B-A*Yr[1]))

        return f(y)
        



