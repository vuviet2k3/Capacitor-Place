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
    Sense
)
PATH_PY = os.path.dirname(__file__)
PATH_RESULT = os.path.join(PATH_PY, 'result')

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
            self.pload = [p / self.s_base / 1000 for p in bus_df['Pload[kW]'].tolist()]
            self.qload = [q / self.s_base / 1000 for q in bus_df['Qload[kVAr]'].tolist()]

            self.type_load = bus_df['Type'].tolist()
            self.res_load = bus_df.loc[bus_df["Type"] == "Residential", "ID"].tolist()
            self.ind_load = bus_df.loc[bus_df["Type"] == "Industrial", "ID"].tolist()
            self.com_load = bus_df.loc[bus_df["Type"] == "Commercial", "ID"].tolist()

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

            self.solver = config['solver']
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
        self.RES_LOAD = Set(
            self.CONTAINER,
            name='RE_LOAD',
            domain=[self.BUS],
            records=self.res_load
        )
        self.IND_LOAD = Set(
            self.CONTAINER,
            name='IND_LOAD',
            domain=[self.BUS],
            records=self.ind_load
        )
        #
        self.COM_LOAD = Set(
            self.CONTAINER,
            name='COM_LOAD',
            domain=[self.BUS],
            records=self.com_load
        )
        #
        self.BUS_attr = Set(
            self.CONTAINER, 
            name='BUS_attr',
            records=['PL', 'QL']
        )
        self.PRF_attr = Set(
            self.CONTAINER,
            name='PRF_attr',
            records=['Residential', 'Industrial', 'Commercial']
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
            domain=[self.BUS],
            records=[self.id_slack],
            description='Tap bus nguon'
        )
        #
        self.BRN = Set(
            self.CONTAINER,
            name='BRN',
            domain=[self.BUS, self.NODE],
            records=list(zip(self.f_bus, self.t_bus)),
            description='Tap nhanh duong day'
        )
        self.BRN_attr = Set(
            self.CONTAINER,
            name='BRN_attr',
            records=['R', 'X', 'RATE']
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
        self.CAP_attr = Set(
            self.CONTAINER,
            name='CAP_attr',
            records=['Qc', 'Cost']
        )

    # define Parameter
    def define_Parameter(self):
        ## grid data
        self.BrnData = Parameter(
            self.CONTAINER,
            name='BrnData',
            domain=[self.BUS, self.NODE, self.BRN_attr],
            records=[
                (f, t, 'R', r) for f, t, r in zip(self.f_bus, self.t_bus, self.R_brn)
            ] + [
                (f, t, 'X', x) for f, t, x in zip(self.f_bus, self.t_bus, self.X_brn)
            ] + [
                (f, t, 'RATE', rate) for f, t, rate in zip(self.f_bus, self.t_bus, self.rateA)
            ],
            description='Thong so nhanh duong day (R, X, RATE)'
        )
        #
        self.BusData = Parameter(
            self.CONTAINER,
            name='BusData',
            domain=[self.BUS, self.BUS_attr],
            records= [
                (f, 'PL', p) for (f, p) in zip(self.id_bus, self.pload)
            ] + [
                (f, 'QL', q) for f, q in zip(self.id_bus, self.qload)
            ],
            description='Thong so bus (PL, QL)'
        )
        #
        self.PrfData =Parameter(
            self.CONTAINER,
            name='PrfData',
            domain=[self.TIME, self.PRF_attr],
            records=[
                (t, 'Residential', prf) for t, prf in zip(self.time, self.res_prf) 
            ] + [
                (t, 'Industrial', prf) for t, prf in zip(self.time, self.ind_prf) 
            ] + [
                (t, 'Commercial', prf) for t, prf in zip(self.time, self.com_prf) 
            ],
            description='Thong so tai theo thoi gian'
        )
        #
        self.CapData = Parameter(
            self.CONTAINER,
            name='CapData',
            domain=[self.CAP, self.CAP_attr],
            records=[
                (cap, 'Qc', q) for cap, q in zip(self.id_cap, self.Q_cap)
            ] + [
                (cap, 'Cost', cost) for cap, cost in zip(self.id_cap, self.cost_cap)
            ],
            description='Thong so cap (Qc, Cost)'
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

    def define_Variable(self):
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
        # self.Ibrn_sqr.up[self.BUS, self.NODE, self.TIME].where[self.BRN[self.BUS, self.NODE]] = self.BrnData[self.BUS, self.NODE, 'RATE']**2
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
        # 
        self.PD = Variable(
            self.CONTAINER,
            name='PD',
            domain=[self.BUS, self.TIME]
        )
        #
        self.QD = Variable(
            self.CONTAINER,
            name='QD',
            domain=[self.BUS, self.TIME]
        )
        # obj variable
        self.OBJ = Variable(
            self.CONTAINER,
            name='OBJ',
            type="free"
        )

    # define equation
    # def define_Equation(self):
    def define_Equation(self):
        self.Eqs_PL = Equation(
            self.CONTAINER, 
            name='Eqs_PL',
            domain=[self.BUS, self.TIME]
        )
        self.Eqs_QL = Equation(
            self.CONTAINER,
            name='Eqs_QL',
            domain=[self.BUS, self.TIME]
        )

        if self.typeload == 'All':
            self.Eqs_PL[self.BUS, self.TIME] = (
                self.PD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'PL'] * self.PrfData[self.TIME, 'Residential']).where[self.RES_LOAD[self.BUS]]
                + (self.BusData[self.BUS, 'PL'] * self.PrfData[self.TIME, 'Industrial']).where[self.IND_LOAD[self.BUS]]
                + (self.BusData[self.BUS, 'PL'] * self.PrfData[self.TIME, 'Commercial']).where[self.COM_LOAD[self.BUS]]
            )
            self.Eqs_QL[self.BUS, self.TIME] = (
                self.QD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'QL'] * self.PrfData[self.TIME, 'Residential']).where[self.RES_LOAD[self.BUS]]
                + (self.BusData[self.BUS, 'QL'] * self.PrfData[self.TIME, 'Industrial']).where[self.IND_LOAD[self.BUS]]
                + (self.BusData[self.BUS, 'QL'] * self.PrfData[self.TIME, 'Commercial']).where[self.COM_LOAD[self.BUS]]
            )
        
        elif self.typeload == 'Residential':
            self.Eqs_PL[self.BUS, self.TIME] = (
                self.PD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'PL'] * self.PrfData[self.TIME, 'Residential']).where[self.RES_LOAD[self.BUS]]
            )
            self.Eqs_QL[self.BUS, self.TIME] = (
                self.QD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'QL'] * self.PrfData[self.TIME, 'Residential']).where[self.RES_LOAD[self.BUS]]
            )
        
        elif self.typeload == 'Industrial':
            self.Eqs_PL[self.BUS, self.TIME] = (
                self.PD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'PL'] * self.PrfData[self.TIME, 'Industrial']).where[self.IND_LOAD[self.BUS]]
            )
            self.Eqs_QL[self.BUS, self.TIME] = (
                self.QD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'QL'] * self.PrfData[self.TIME, 'Industrial']).where[self.IND_LOAD[self.BUS]]
            )
        
        elif self.typeload == 'Commercial':
            self.Eqs_PL[self.BUS, self.TIME] = (
                self.PD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'PL'] * self.PrfData[self.TIME, 'Commercial']).where[self.COM_LOAD[self.BUS]]
            )
            self.Eqs_QL[self.BUS, self.TIME] = (
                self.QD[self.BUS, self.TIME] ==
                (self.BusData[self.BUS, 'QL'] * self.PrfData[self.TIME, 'Commercial']).where[self.COM_LOAD[self.BUS]]
            )
        
        # eqs (13) - (14)
        self.Eqs131 = Equation(
            self.CONTAINER,
            name='eqs131',
            domain=[self.BUS, self.TIME]
        )
        self.Eqs132 = Equation(
            self.CONTAINER,
            name='eqs132',
            domain=[self.SLACK, self.TIME]
        )
        #
        self.Eqs141 = Equation(
            self.CONTAINER,
            name='eqs141',
            domain=[self.BUS, self.TIME]
        )
        self.Eqs142 = Equation(
            self.CONTAINER,
            name='eqs142',
            domain=[self.SLACK, self.TIME]
        )

        self.Eqs131[self.BUS, self.TIME].where[~self.SLACK[self.BUS]] = (
            Sum(
                self.NODE.where[self.BRN[self.NODE, self.BUS] & (self.NODE.ord < self.BUS.ord)],
                self.Pbrn[self.NODE, self.BUS, self.TIME] - 
                self.BrnData[self.NODE, self.BUS, 'R'] * self.Ibrn_sqr[self.NODE, self.BUS, self.TIME]
            ) == 
            Sum(
                self.NODE.where[self.BRN[self.BUS, self.NODE] & (self.BUS.ord < self.NODE.ord)],
                self.Pbrn[self.BUS, self.NODE, self.TIME]
            ) + self.PD[self.BUS, self.TIME]
        )
        
        self.Eqs132[self.SLACK, self.TIME] = (
            Sum(
                self.NODE.where[self.BRN[self.SLACK, self.NODE]],
                self.Pbrn[self.SLACK, self.NODE, self.TIME]
            ) == self.Pgen[self.SLACK, self.TIME]
        )
        #
        self.Eqs141[self.BUS, self.TIME].where[~self.SLACK[self.BUS]] = (
            Sum(
                self.NODE.where[self.BRN[self.NODE, self.BUS] & (self.NODE.ord < self.BUS.ord)],
                self.Qbrn[self.NODE, self.BUS, self.TIME] - 
                self.BrnData[self.NODE, self.BUS, 'X'] * self.Ibrn_sqr[self.NODE, self.BUS, self.TIME]
            ) == 
            Sum(
                self.NODE.where[self.BRN[self.BUS, self.NODE] & (self.BUS.ord < self.NODE.ord)],
                self.Qbrn[self.BUS, self.NODE, self.TIME]
            ) + self.PD[self.BUS, self.TIME] - self.Qcap[self.BUS]
        )
        #
        self.Eqs142[self.SLACK, self.TIME] = (
            Sum(
                self.NODE.where[self.BRN[self.SLACK, self.NODE]],
                self.Qbrn[self.SLACK, self.NODE, self.TIME]
            ) == self.Qgen[self.SLACK, self.TIME]
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
            2 * (self.BrnData[self.BUS, self.NODE, 'R'] * self.Pbrn[self.BUS, self.NODE, self.TIME] +
                 self.BrnData[self.BUS, self.NODE, 'X'] * self.Qbrn[self.BUS, self.NODE, self.TIME]) +
            (self.BrnData[self.BUS, self.NODE, 'R']**2 + self.BrnData[self.BUS, self.NODE, 'X']**2) * self.Ibrn_sqr[self.BUS, self.NODE, self.TIME]
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
            self.Qcap[self.BUS] == Sum(self.CAP, self.CapData[self.CAP, 'Qc'] * self.Zcap[self.BUS, self.CAP])
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
    # define OBJ
    def define_Obj(self):
        self.Eqs_Obj = Equation(
            self.CONTAINER,
            name='Eqs_Obj',
        )

        self.Eqs_Obj[...] = (
            self.OBJ == 
            365 * Sum(
                (self.BUS, self.NODE, self.TIME),
                (self.BrnData[self.BUS, self.NODE, 'R'] * 
                self.Ibrn_sqr[self.BUS, self.NODE, self.TIME] *
                self.SBASE.toValue() * self.COSTA.toValue()
                ).where[self.BRN[self.BUS, self.NODE] & (self.BUS.ord < self.NODE.ord)]
            )
            + Sum(
                (self.BUS, self.CAP),
                self.Zcap[self.BUS, self.CAP] * 
                self.CapData[self.CAP, 'Cost'] *
                self.CapData[self.CAP, 'Qc'] * 
                self.SBASE.toValue() *
                self.R.toValue() * (1 + self.R.toValue())**self.MCAP.toValue() / 
                ((1 + self.R.toValue())**self.MCAP.toValue() - 1)
            )
        )
    # define Options
    def define_Options(self):
        self.opts = Options()
        self.opts.equation_listing_limit = 10000000
        self.opts.absolute_optimality_gap = 0.0          # optca
        self.opts.relative_optimality_gap = 0.0          # optcr
        self.opts.miqcp = self.solver
        self.opts.listing_file = os.path.join(PATH_RESULT, 'misocp2.lst')

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
        self.MODEL.toGams(os.path.join(PATH_RESULT,'misocp2.gms'))

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

    print("\n=== Xuất kết quả ===")
    
    # Xử lý file đang mở
    output_file = os.path.join(PATH_RESULT, "result.xlsx")
    if os.path.exists(output_file):
        try:
            os.remove(output_file)
        except PermissionError:
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            output_file = f"result_{timestamp}.xlsx"
            print(f"File result.xlsx đang mở, lưu vào: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Voltage - Lấy tất cả cột
        df = opt.Usqr.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Voltage", index=False)
            print(f"  Voltage: {len(df)} rows, {len(df.columns)} cols")
        
        # Active Power
        df = opt.Pbrn.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Pbrn", index=False)
            print(f"  Pbrn: {len(df)} rows, {len(df.columns)} cols")
        
        # Reactive Power
        df = opt.Qbrn.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Qbrn", index=False)
            print(f"  Qbrn: {len(df)} rows, {len(df.columns)} cols")
        
        # Capacitor Q
        df = opt.Qcap.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Qcap", index=False)
            print(f"  Qcap: {len(df)} rows, {len(df.columns)} cols")
        
        # Capacitor binary
        df = opt.Zcap.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Zcap", index=False)
            # Lọc những vị trí có tụ (cột 'level' > 0.5)
            df_installed = df[df.iloc[:, 2] > 0.5]  # Cột thứ 3 là 'level'
            df_installed.to_excel(writer, sheet_name="Cap_Installed", index=False)
            print(f"  Zcap: {len(df)} rows, Installed: {len(df_installed)}")
        
        # Current squared
        df = opt.Ibrn_sqr.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Ibrn_sqr", index=False)
            print(f"  Ibrn_sqr: {len(df)} rows, {len(df.columns)} cols")
        
        # P Generation
        df = opt.Pgen.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Pgen", index=False)
            print(f"  Pgen: {len(df)} rows, {len(df.columns)} cols")
        
        # Q Generation
        df = opt.Qgen.records
        if not df.empty:
            df.to_excel(writer, sheet_name="Qgen", index=False)
            print(f"  Qgen: {len(df)} rows, {len(df.columns)} cols")
        
        # Load P
        df = opt.PD.records
        if not df.empty:
            df.to_excel(writer, sheet_name="PD", index=False)
            print(f"  PD: {len(df)} rows, {len(df.columns)} cols")
        
        # Load Q
        df = opt.QD.records
        if not df.empty:
            df.to_excel(writer, sheet_name="QD", index=False)
            print(f"  QD: {len(df)} rows, {len(df.columns)} cols")
        
        # Objective
        obj_val = opt.OBJ.toValue()
        pd.DataFrame({
            "Objective": [obj_val],
            "Description": ["Total Cost ($/year)"]
        }).to_excel(writer, sheet_name="Objective", index=False)
        print(f"  Objective: {obj_val:.2f}")
    
    print(f"\nKết quả đã lưu: {output_file}")
    print(f"Tổng chi phí: ${obj_val:,.2f}")



if __name__=='__main__':
    main()
    