from start_powerfactory import start_powerfactory
import os
import pandas as pd

#variable pre-allocation
Bus_col = []
Skss_col = []
Ikss_col = []
Vbus_col = []

def set_load_flow_settings(ldf_com_obj, load_scaling, generation_scaling):

    ldf_com_obj.SetAttribute('scLoadFac', load_scaling)          # set load scaling factor
    ldf_com_obj.SetAttribute('scGenFac', generation_scaling)     # set generation scaling factor

    ldf_com_obj.SetAttribute('iopt_net', 0)  # AC Load Flow, balanced, positive sequence
    ldf_com_obj.SetAttribute('iopt_at', 0)  # Automatic tap adjustment of transformers
    ldf_com_obj.SetAttribute('iopt_pq', 0)  # Consider Voltage Dependency of Loads

    return


file = 'PF QDS solved'

####################### start PowerFactory ###################################

app, study_case_obj, ldf, o_ElmNet = start_powerfactory(file)
app.Show()

##################### parameterize a load flow & short circuit simulation, execute it and export results ##################################

terminals = app.GetCalcRelevantObjects("*.ElmTerm")
Pvs = app.GetCalcRelevantObjects("*.ElmPVsys")

#########################################################################
###### short circuit #######################
#########################################################################

oSC = app.GetFromStudyCase('ComShc') #get SC object
setattr(oSC, 'iopt_mde',
        0)  # SC calculation method: 0 = VDE 0102 Part 0 / DIN EN 60909-0 (six other methods avail. -> check in SC calculation object in PF)
setattr(oSC, 'iopt_shc',
        '3psc')  # SC type: 3psc = 3-phase, 2psc = 2-phase, spgf = single-phase-to-ground, 2pgf = 2-phase-to-ground (more options -> check in SC calculation object in PF)
setattr(oSC, 'Rf', 0)  # (Ohm) fault resistance
setattr(oSC, 'Xf', 0)  # (Ohm) fault reactance
setattr(oSC, 'iopt_allbus',
        0)  # fault location option: 0 = User Selection, 1 = Busbars and Junction Nodes, 2 = All Busbars

o_PVsys = Pvs[0]
PV_cubicle = o_PVsys.bus1  # elements are connected to terminals via cubicles in powerfactory
o_ElmTerm = PV_cubicle.cterm # PV bus

setattr(oSC, 'shcobj', o_ElmTerm)  # fault location -> bus/terminal object
oSC.Execute()  # execute load flow

# get result variables of interest
Skss = getattr(o_ElmTerm, 'm:Skss')  # (kVA) short-circuit power
Ikss = getattr(o_ElmTerm, 'm:Ikss')  # (kA) initial symmetrical short-circuit current
Vbus = getattr(o_ElmTerm, 'uknom')  # (kV) voltage level of bus

# put together important data
Bus_col.append(getattr(o_ElmTerm, 'loc_name'))  # bus name
Skss_col.append(Skss)  # (MVA) short-circuit power
Ikss_col.append(Ikss)  # (kA) initial symmetrical short-circuit current
Vbus_col.append(Vbus)  # (kV) voltage level of bus

#Results: bus names, SC power, SC current, bus voltage
results = pd.DataFrame()
results['name'] = Bus_col
results['Skss'] = Skss_col
results['Ikss'] = Ikss_col
results['Vbus'] = Vbus_col

results.to_csv(os.getcwd() + '\\results short circuit calc.csv', header=True, sep=';', decimal='.',
                       float_format='%.' + '%sf' % 2)

#########################################################################
##### load flow #########################################################################
#########################################################################

set_load_flow_settings(ldf, 100, 75)  # Set default load flow settings

ldf.Execute()

## write to result file
elmres = app.GetFromStudyCase('Results.ElmRes') #results element

for terminal in terminals:
    for var in ['m:u', 'm:phiu', 'm:Pflow', 'm:Qflow']:
        elmres.AddVariable(terminal, var)
elmres.Write()
elmres.FinishWriting()

# define result export
comres = app.GetFromStudyCase('ComRes') # export of a resultfile can be done using the ComRes command
comres.iopt_sep = 0  # to use the system seperator
comres.iopt_honly = 0  # to export data and not only the header
comres.iopt_tsel = 1    # export only selected variables
comres.iopt_exp = 6  # to export as csv

comres.pResult = elmres #assign the results file to the export function
comres.f_name = os.getcwd() + '\\results load flow calc.csv'
comres.Execute()
##################### parameterize a QDS simulation, execute it and export results ##################################

def set_QDS_settings(app, study_case_obj, t_start, t_end, step_unit=1, balanced=0, step_size=None, marker=None):
    if step_size is None:
        step_size = config.step_size

    local_machine_tz = config.local_machine_tz

    if study_case_obj.SearchObject('*.ComStatsim') is None:
        qds = app.GetFromStudyCase("ComStatsim")
        qds.Execute()

    qds_com_obj = study_case_obj.SearchObject('*.ComStatsim')
    qds_com_obj.SetAttribute('iopt_net', balanced)  # AC Load Flow, balanced, positive sequence
    qds_com_obj.SetAttribute('calcPeriod', 4)  # User defined time range

    qds_com_obj.SetAttribute('startTime', int(
        (pd.DatetimeIndex([(t_start - t_start.tz_convert(local_machine_tz).utcoffset())]).astype(np.int64) // 10 ** 9)[
            0]))
    qds_com_obj.SetAttribute('endTime', int(
        (pd.DatetimeIndex([(t_end - t_end.tz_convert(local_machine_tz).utcoffset())]).astype(np.int64) // 10 ** 9)[
            0]))

    qds_com_obj.SetAttribute('stepSize', step_size)
    qds_com_obj.SetAttribute('stepUnit', step_unit)  # Seconds, Minutes ....

    if config.parallel_computing == True:
        qds_com_obj.SetAttribute('iEnableParal', config.parallel_computing)  # settings for parallel computation

        user = app.GetCurrentUser()
        settings = user.SearchObject(r'Set\Def\Settings')
        settings.SetAttribute('iActUseCore', 1)
        settings.SetAttribute('iCoreInput', config.cores)

    if config.just_voltages:  # define result variables of interest; result file is generated accordingly
        result_variables = {
            'ElmTerm': [
                'm:u',  # voltage at terminal
            ],
        }
    elif config.deeplearning:
        result_variables = {
            'ElmTerm': [
                'm:u',  # voltage at terminal
                'm:Pgen',  # active power generated at terminal
                'm:Qgen',  # reactive power generated at terminal
            ],
        }
    elif config.sim_setting in ['stmk']:
        result_variables = {
            'ElmTerm': [
                'm:u',
                'm:Pflow',
                'm:Qflow'
            ],
            'ElmLne': {
                'm:P:bus2',
                'm:Q:bus2',
        }}

    elif marker == 'load_estimation_training_data':
        result_variables = {
            'ElmTerm': [
                'm:u',
                'm:phiu',
                'm:Pflow',
                'm:Qflow'
            ],
            'ElmLod': [
                'm:P:bus1',
                'm:Q:bus1'
            ],
            'ElmPvsys': [
                'm:P:bus1',
                'm:Q:bus1'
            ]
        }
    else:
        result_variables = m.mapping()

    result = pf.app.GetFromStudyCase('ComStatsim').results  # should be language independent
    """if config.system_language == 0:
        #result = pf._resolve_result_object('Quasi-Dynamic Simulation AC')  # get result file
        result = pf.app.GetFromStudyCase('ComStatsim').presults #should be language independent
    else:
        if config.system_language == 1:
            result = pf.app.GetFromStudyCase('ComStatsim').presults #should be language independent
            #result = pf._resolve_result_object('Quasi-Dynamische Simulation AC')  # get result file"""

    pf.set_vars_of_result_obj(result,
                              result_variables=result_variables)

    return result

