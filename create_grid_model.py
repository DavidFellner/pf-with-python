from start_powerfactory import start_powerfactory
import os
#parent_dir = os.path.dirname(os.getcwd())
#path = parent_dir + "\\Block 3 'DigSilent Power Factory & Python basics, automated model creation, and assignment of profiles'"
#file = path + "\\PF LF&SC base.pfd"

file = 'PF LF&SC base'

####################### start PowerFactory ###################################

app, study_case_obj, ldf, o_ElmNet = start_powerfactory(file)
app.Show()

#################### create grid model ####################################

#create terminal + set attribute for rated voltage
o_ElmTerm = o_ElmNet.CreateObject('ElmTerm', '110 kV SS')
o_ElmTerm.SetAttribute('uknom', 110)

#create connection point at terminal ('cubicle')
o_StaCubic = o_ElmTerm.CreateObject('StaCubic', 'Field_1')

#create external grid + link it to terminal
o_ElmXNet = o_ElmNet.CreateObject('ElmXNet', 'External grid')
o_ElmXNet.SetAttribute('snss', 1000)
o_ElmXNet.SetAttribute('rntxn', 1/6)
o_ElmXNet.SetAttribute('bus1', o_StaCubic)

#create transformer
o_ElmTr2 = o_ElmNet.CreateObject('ElmTr2', 'HV_MV_transformer')
type = [i for i in app.GetProjectFolder('equip').GetContents() if i.loc_name == '110 kV/ 20kV'][0] #get type from library
o_ElmTr2.SetAttribute('typ_id', type)

#connect transformer to both sides
o_StaCubic_hv = o_ElmTerm.CreateObject('StaCubic', 'Field_2')
o_ElmTr2.SetAttribute('bushv', o_StaCubic_hv)

o_StaCubic_lv = app.GetCalcRelevantObjects('20 kV SS.ElmTerm')[0].CreateObject('StaCubic', 'Transformers_lv_side')
o_ElmTr2.SetAttribute('buslv', o_StaCubic_lv)

#create line and connect it
o_ElmLne = o_ElmNet.CreateObject('ElmLne', 'L Stat2_AP PV')
type = [i for i in app.GetProjectFolder('equip').GetContents() if i.loc_name == 'AL/St 95/15'][0]
o_ElmLne.SetAttribute('typ_id', type)
o_ElmLne.SetAttribute('dline', 11)

o_StaCubic_1 = app.GetCalcRelevantObjects('Stat2.ElmTerm')[0].CreateObject('StaCubic', 'line_side1')
o_StaCubic_2 = app.GetCalcRelevantObjects('VP PV.ElmTerm')[0].CreateObject('StaCubic', 'line_side 2')
o_ElmLne.SetAttribute('bus1', o_StaCubic_1)
o_ElmLne.SetAttribute('bus2', o_StaCubic_2)

#create load and connect it
o_ElmLod = o_ElmNet.CreateObject('ElmLod', 'Load 2')
o_ElmLod.SetAttribute('plini', 1)
o_ElmLod.SetAttribute('coslini', 0.9)

o_StaCubic = app.GetCalcRelevantObjects('Stat2.ElmTerm')[0].CreateObject('StaCubic', 'load')
o_ElmLod.SetAttribute('bus1', o_StaCubic)

#create PV generator and connect it
o_ElmPVsys = o_ElmNet.CreateObject('ElmPVsys', 'PV generator')
o_ElmPVsys.SetAttribute('sgn', 2000)
o_ElmPVsys.SetAttribute('cosn', 0.8)
o_ElmPVsys.SetAttribute('pgini', 1.5)
o_ElmPVsys.SetAttribute('cosgini', 1)

o_StaCubic = app.GetCalcRelevantObjects('WR_PV.ElmTerm')[0].CreateObject('StaCubic', 'PV')
o_ElmPVsys.SetAttribute('bus1', o_StaCubic)

#create load/generation profile and assign it
charFold = app.GetProjectFolder('chars') #get folder, where characteristics are normally stored
char = charFold.CreateObject('ChaTime','PV_char')

char.SetAttribute('source', 1) #data from file
char.SetAttribute('ciopt_stamp', 1) #time stamped data
char.SetAttribute('timeformat', 'DD.MM.YYYY hh:mm') #time format in file
char.SetAttribute('f_name', os.getcwd() + '\\Profiles.csv') #set file path
char.SetAttribute('iopt_sep', 0) #use system separators?
char.SetAttribute('dec_Sep', '.') #use . as seperator
char.SetAttribute('datacol', 3) #use 3rd column
char.SetAttribute('usage', 2) #use the absolute values in the profile

char_ref = o_ElmPVsys.CreateObject('ChaRef','pgini') #create a characteristic reference object for a certain parameter
char_ref.typ_id = char #link the characteristic to the reference object

app.Hide()