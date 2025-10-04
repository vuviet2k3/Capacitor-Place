import os, sys
import pandas as pd
import json 
from gamspy import (
    Container,
    Set,
    Alias,
    Parameter,
    Variable,
    Equation,
    Sum,
    Options,
    Model,
    Problem,
    Sense,
)

class GetData:
    def __init__(self, xlsx_file=None, json_file=None):
        self.xlsx_file = xlsx_file or 'ieee33.xlsx'
        self.json_file = json_file or 'config.json'

    def get_xlsx(self):
        try:
            ## bus data
            bus_df = pd.read_excel(self.xlsx_file, sheet_name='bus', header=1)
            #
            self.id_bus = bus_df['ID'].tolist()
            self.name_bus = bus_df['Name'].tolist()

            ## load data
            load = bus_df[bus_df['Type'].notna()]
            self.id_load = load['ID'].tolist()
            self.pload = [p / self.s_base / 1000 for p in load['Pload[kW]'].tolist()]
            self.qload = [q / self.s_base / 1000 for q in load['Qload[kVAr]'].tolist()]
            self.type_load = load['Type'].tolist()

            # convert slack bus code = 3 and vsch (p.u)
            slack = bus_df[bus_df['Code'] == 3]
            self.id_slack = int(slack['ID'][0])
            self.u_slack = slack['Vsch[pu]'][0]

            ### line data
            line_df = pd.read_excel(self.xlsx_file, sheet_name='line', header=1)
            #
            self.id_line = line_df['ID'].tolist()
            self.f_bus = line_df['FromBus'].tolist()
            self.t_bus = line_df['ToBus'].tolist()
            self.R_brn = [r / self.z_base for r in line_df['R[Ohm]'].tolist()]
            self.X_brn = [x / self.z_base for x in line_df['X[Ohm]'].tolist()]
            self.rateA = [rate / self.i_base for rate in line_df['rateA[kA]'].tolist()]
   
            # load profile data
            loadprofile_df = pd.read_excel(self.xlsx_file, sheet_name='loadprofile', header=1)
            #
            self.time = loadprofile_df['Time'].tolist()
            self.res_prf = loadprofile_df['Residential'].tolist()
            self.com_prf = loadprofile_df['Commercial'].tolist()
            self.ind_prf = loadprofile_df['Industrial'].tolist()
            
            # capacitor data
            capacitor_df = pd.read_excel(self.xlsx_file, sheet_name='capacitor', header=1)
            #
            self.id_cap = capacitor_df['ID'].tolist()
            self.type_cap = capacitor_df['Type'].tolist()
            self.Q_cap = [q / self.s_base / 1000 for q  in capacitor_df['Size[kVAr]'].tolist()]
            self.cost_cap = [cost * 1000 for cost in capacitor_df['Cost[$/kVAr]'].tolist()]

        #
        except Exception as e:
            print(f'Loi doc file.xlsx: {e}')

    def get_json(self):
        try:
            with open(self.json_file, 'r', encoding='utf-8') as f:
                config = json.load(f)

            # data
            self.typeload = config['data']['type_load']

            # base data
            self.s_base = config['base']['s_base']
            self.u_base = config['base']['u_base']
            self.z_base = self.u_base**2 / self.s_base
            self.i_base = self.s_base / self.u_base / 3**(1/2)

            # voltage limit
            self.u_min = config['volt_limit']['volt_lower']
            self.u_max = config['volt_limit']['volt_upper']

            # economic data
            self.cost_A = config['economic_parameters']['c_delta_a']    # gia dien nang
            self.r = config['economic_parameters']['r']     # he so chiet khau
            self.M = config['economic_parameters']['M']     # tuoi tho tu bu ngang
            self.Y = config['economic_parameters']['Y']     # tong so vi tri dat tu bu ngang

        #
        except Exception as e:
            print(f'Loi doc file.json: {e}')

    

class MISOCP(GetData):
    def __init__(self, xlsx_file=None, json_file=None):
        super().__init__(xlsx_file, json_file)
        self.get_json()
        self.get_xlsx()
        self.CONTAINER = Container()

    # difine Set 
    def define_Set(self):
        self.BUS = Set(
            self.CONTAINER,
            name='BUS',
            records=self.id_bus,
            description='Tap bus'
        )
        #
        self.NODE = Alias(
            self.CONTAINER,
            name='NODE',
            alias_with=self.BUS,
        )
        #
        self.SLACK = Set(
            self.CONTAINER,
            name='SLACK',
            records=[self.id_slack],
            description='Tap bus nguon'
        )
        #
        self.LOAD = Set(
            self.CONTAINER,
            name='LOAD',
            records=self.id_load,
            description='Tap phu tai'
        )
        #
        self.BRN = Set(
            self.CONTAINER,
            name='BRN',
            domain=[self.BUS, self.NODE],
            records=list(zip(self.f_bus, self.t_bus)),
            description='Tap nhanh duong day'
        )
        self.LINE = Set(
            self.CONTAINER,
            name='LINE',
            records=self.id_line,
            description='Tap ID duong day'
        )
        #
        self.TIME = Set(
            self.CONTAINER,
            name='TIME',
            records=self.time,
            description='Tap thoi gian'
        )
        #
        self.CAP = Set(
            self.CONTAINER,
            name='CAP',
            records=self.id_cap,
            description='Tap du lieu Capacitor'
        )

    # define Parameter
    def define_Parameter(self):
        ## grid data
        self.RBRN = Parameter(
            self.CONTAINER, 
            name='RBRN',
            domain=[self.BUS, self.NODE],
            records=list(zip(self.f_bus, self.t_bus, self.R_brn)),
            description='Khai bao gia tri R_ij'
        )
        self.XBRN = Parameter(
            self.CONTAINER, 
            name='XBRN',
            domain=[self.BUS, self.NODE],
            records=list(zip(self.f_bus, self.t_bus, self.X_brn)),
            description='Khai bao gia tri X_ij'
        )
        #
        self.RATEBRN = Parameter(
            self.CONTAINER,
            name='RATEBRN',
            domain=[self.BUS, self.NODE],
            records=list(zip(self.f_bus, self.t_bus, self.rateA))
        )
        #
        self.PLOAD = Parameter(
            self.CONTAINER,
            name='PLOAD',
            domain=[self.LOAD, self.TIME],
            description='Khai bao gia tri Pload'
        )
        #
        self.QLOAD = Parameter(
            self.CONTAINER,
            name='QLOAD',
            domain=[self.LOAD, self.TIME],
            description='Khai bao gia tri Qload'
        )
        
        prf_map = {
            "Residential": self.res_prf,
            "Commercial": self.com_prf,
            "Industrial": self.ind_prf,
        }
        if self.typeload == 'All':
            buses_to_process = zip(self.id_load, self.pload, self.qload, self.type_load)
        else:
            buses_to_process = [
                (load, pload, qload, load_type) 
                for load, pload, qload, load_type in zip(self.id_load, self.pload, self.qload, self.type_load)
                if load_type == self.typeload
            ]

        for bus, pload, qload, load_type in buses_to_process:
            prf_list = prf_map[load_type]  
            for tdx, time in enumerate(self.time):
                prf = prf_list[tdx]
                self.PLOAD[bus, time] = pload * prf
                self.QLOAD[bus, time] = qload * prf

        # capacitor data
        self.SIZECAP = Parameter(
            self.CONTAINER,
            name='SIZECAP',
            domain=self.CAP,
            records=list(zip(self.id_cap, self.Q_cap)),
            description='Khai bao dung luong tu bu'
        )
        #
        self.COSTCAP = Parameter(
            self.CONTAINER,
            name='COSTCAP',
            domain=self.CAP,
            records=list(zip(self.id_cap, self.cost_cap)),
            description='Khai bao gia tu bu'
        )
        #
        self.YCAP = Parameter(
            self.CONTAINER,
            name='YCAP',
            records=self.Y,
            description='Tong so lap dat tu bu'
        )
        #
        self.MCAP = Parameter(
            self.CONTAINER,
            name='MCAP',
            records=self.M,
            description='Tuoi tho tu bu'
        )

        # economic data
        self.SBASE = Parameter(
            self.CONTAINER,
            name='SBASE',
            records=self.s_base,
            description='Khai bao gia tri Sbase'
        )
        #
        self.UMIN = Parameter(
            self.CONTAINER,
            name='UMIN',
            records=self.u_min,
            description='Khai bao gia tri Umin'
        )
        #
        self.UMAX = Parameter(
            self.CONTAINER,
            name='UMAX',
            records=self.u_max,
            description='Khai bao gia tri Umax'
        )
        #
        self.COSTA = Parameter(
            self.CONTAINER,
            name='COSTA',
            records=self.cost_A,
            description='Khai bao gia ton that dien nang'
        )
        #
        self.R = Parameter(
            self.CONTAINER,
            name='R',
            records=self.r,
            description='Khai bao he so chiet khau'
        )

    # define Variable
    def define_Variable(self):

        # grid variable
        self.Usqr = Variable(
            self.CONTAINER,
            name='Usqr',
            domain=[self.BUS, self.TIME],
            type="positive"
        )
        self.Usqr.lo[self.BUS, self.TIME] = self.UMIN.toValue()**2
        self.Usqr.up[self.BUS, self.TIME] = self.UMAX.toValue()**2
        self.Usqr.fx[self.id_slack, self.TIME] = self.u_slack**2

        #
        self.Ibrn_sqr = Variable(
            self.CONTAINER,
            name='IBRNsqr',
            domain=[self.BUS, self.NODE, self.TIME],
            type="positive"
        )
        for (fbus, tbus, rate) in zip(self.f_bus, self.t_bus, self.rateA):
            self.Ibrn_sqr.up[fbus, tbus, self.TIME] = rate**2

        #
        self.Pbrn = Variable(
            self.CONTAINER, 
            name='Pbrn',
            domain=[self.BUS, self.NODE, self.TIME],
            type="free"
        )
        #
        self.Qbrn = Variable(
            self.CONTAINER, 
            name='Qbrn',
            domain=[self.BUS, self.NODE, self.TIME],
            type="free"
        )
        #
        self.Pgen = Variable(
            self.CONTAINER,
            name='Pgen',
            domain=[self.SLACK, self.TIME],
            type="free"
        )
        #
        self.Qgen = Variable(
            self.CONTAINER,
            name='Qgen',
            domain=[self.SLACK, self.TIME],
            type="free"
        )

        # capacitor variable
        self.Qcap = Variable(
            self.CONTAINER,
            name='Qcap',
            domain=[self.BUS],
            type="free"
        )
        #
        self.Zcap = Variable(
            self.CONTAINER,
            name='Zcap',
            domain=[self.BUS, self.CAP],
            type="binary"
        )

        # obj variable
        self.OBJ = Variable(
            self.CONTAINER,
            name='OBJ',
            type="free"
        )

        return True
    
    def get_parent(self, parent):
        return [fbus for (fbus, tbus) in zip(self.f_bus, self.t_bus) if tbus == parent]
    
    def get_child(self, child):
        return [tbus for (fbus, tbus) in zip(self.f_bus, self.t_bus) if fbus == child]
    
    # define equation
    def define_Equation(self):

        # eqs (13) - (14)
        self.Eqs13 = Equation(
            self.CONTAINER,
            name='eqs13',
            domain=[self.BUS, self.TIME]
        )
        #
        self.Eqs14 = Equation(
            self.CONTAINER,
            name='eqs14',
            domain=[self.BUS, self.TIME]
        )

        for fbus in self.id_bus:
            for t in self.time:
                self.Eqs13[fbus, t] = (
                    (self.Pgen[fbus, t] if fbus == self.id_slack else 0) +
                    sum(
                        (self.Pbrn[parent, fbus, t] - self.RBRN[parent, fbus] * self.Ibrn_sqr[parent, fbus, t]) 
                        for parent in self.get_parent(fbus)
                        )
                    ==
                    sum(self.Pbrn[fbus, child, t] for child in self.get_child(fbus) if child != fbus) +
                    (self.PLOAD[fbus, t] if fbus in self.id_load else 0)
                )
                #
                self.Eqs14[fbus, t] = (
                    (self.Qgen[fbus, t] if fbus == self.id_slack else 0) +
                    sum(
                        (self.Qbrn[parent, fbus, t] - self.XBRN[parent, fbus] * self.Ibrn_sqr[parent, fbus, t])
                        for parent in self.get_parent(fbus)
                    ) +
                    self.Qcap[fbus]
                    ==
                    sum(self.Qbrn[fbus, child, t] for child in self.get_child(fbus) if child != fbus) +
                    (self.QLOAD[fbus, t] if fbus in self.id_load else 0)
                )

        # eqs (15) - (16)
        self.Eqs15 = Equation(
            self.CONTAINER,
            name='Eqs15',
            domain=[self.BUS, self.NODE, self.TIME]
        )
        #
        self.Eqs16 = Equation(
            self.CONTAINER,
            name='Eqs16',
            domain=[self.BUS, self.NODE, self.TIME]
        )
        self.Eqs15[self.BUS, self.NODE, self.TIME].where[self.BRN[self.BUS, self.NODE]] = (
            self.Usqr[self.BUS, self.TIME] - self.Usqr[self.NODE, self.TIME] -
            2 * (self.RBRN[self.BUS, self.NODE] * self.Pbrn[self.BUS, self.NODE, self.TIME] +
                 self.XBRN[self.BUS, self.NODE] * self.Qbrn[self.BUS, self.NODE, self.TIME]) +
            (self.RBRN[self.BUS, self.NODE]**2 + self.XBRN[self.BUS, self.NODE]**2) * self.Ibrn_sqr[self.BUS, self.NODE, self.TIME]
            == 0
        )
        #
        self.Eqs16[self.BUS, self.NODE, self.TIME].where[self.BRN[self.BUS, self.NODE]] = (
            self.Ibrn_sqr[self.BUS, self.NODE, self.TIME] * self.Usqr[self.BUS, self.TIME]
            >=
            self.Pbrn[self.BUS, self.NODE, self.TIME]**2 + self.Qbrn[self.BUS, self.NODE, self.TIME]**2
        )

        # eqs (20)
        self.Eqs20 = Equation(
            self.CONTAINER,
            name='Eqs20',
            domain=self.BUS
        )
        self.Eqs20[self.BUS] = (
            self.Qcap[self.BUS] == Sum(self.CAP, self.SIZECAP[self.CAP] * self.Zcap[self.BUS, self.CAP])
        )

        # eqs (9) - (10)
        self.Eqs9 = Equation(
            self.CONTAINER,
            name='Eqs9',
        )
        self.Eqs9[...] = (
            Sum((self.BUS, self.CAP), self.Zcap[self.BUS, self.CAP]) <= self.YCAP.toValue()
        )
        #
        self.Eqs10 = Equation(
            self.CONTAINER,
            name='Eqs10',
            domain=self.BUS
        )
        self.Eqs10[self.BUS] = (
            Sum(self.CAP, self.Zcap[self.BUS, self.CAP]) <= 1
        )

        return True
    
    # define OBJ
    def define_Obj(self):
        self.Eqs_Obj = Equation(
            self.CONTAINER,
            name='Eqs_Obj',
        )

        self.Eqs_Obj[...] = (
            self.OBJ == 365 * sum(
                # (self.BUS, self.NODE, self.TIME),
                # self.RBRN[self.BUS, self.NODE] *
                # self.Ibrn_sqr[self.BUS, self.NODE, self.TIME] *
                # self.SBASE.toValue() * self.COSTA.toValue()
                self.RBRN[fbus, tbus] *
                self.Ibrn_sqr[fbus, tbus, time] *
                self.SBASE.toValue() * self.COSTA.toValue()
                for (fbus, tbus, time) in zip(self.f_bus, self.t_bus, self.time)
            )
            + Sum(
                (self.BUS, self.CAP),
                self.Zcap[self.BUS, self.CAP] * self.SBASE.toValue() *
                self.COSTCAP[self.CAP] * self.SIZECAP[self.CAP] *
                self.R.toValue() * (1 + self.R.toValue())**self.MCAP.toValue() / 
                ((1 + self.R.toValue())**self.MCAP.toValue() - 1)
            )
        ) 
        
        return True
    
    # define Options
    def define_Options(self):
        self.opts = Options()
        self.opts.equation_listing_limit = 10000000
        self.opts.absolute_optimality_gap = 0.0          # optca
        self.opts.relative_optimality_gap = 0.0          # optcr
        self.opts.miqcp = "cplex" 
        self.opts.listing_file = 'misocp.lst'

        return True
    
    # define Model
    def define_Model(self):
        self.MODEL = Model(
            self.CONTAINER,
            name='Capacitor_Place',
            equations=self.CONTAINER.getEquations(),
            sense=Sense.MIN,
            objective=self.OBJ,
            problem=Problem.MIQCP
        )

        return True
    
    def Solve(self):
        self.MODEL.solve(options=self.opts, output=sys.stdout)

        return True

def main():
    input_xlsx = r"D:\OAEM Lab\CodePy\Capacitor Place\ieee33.xlsx"
    input_json = r"D:\OAEM Lab\CodePy\Capacitor Place\config.json"

    opt = MISOCP(input_xlsx, input_json)
    opt.define_Set()
    opt.define_Parameter()
    opt.define_Variable()
    opt.define_Equation()
    opt.define_Obj()
    opt.define_Options()
    opt.define_Model()
    opt.Solve()

if __name__=='__main__':
    main()
    
