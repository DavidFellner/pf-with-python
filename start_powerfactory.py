import os

def start_powerfactory(file):

    import sys
    os.environ["PATH"] = r"C:\Program Files\DIgSILENT\PowerFactory 2023 SP4;" + os.environ["PATH"]
    sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2023 SP4\Python\3.9")

    import powerfactory


    try:
        app = powerfactory.GetApplicationExt()
    except powerfactory.ExitError as error:
        print(error)
        print('error.code = %d' % error.code)

    app.ActivateProject(file)

    project = app.GetActiveProject()
    study_case_obj = app.GetActiveStudyCase()                                   # Get active study case

    if not study_case_obj.SearchObject('*.ComLdf'):
        study_case_obj.CreateObject('ComLdf', 'load_flow_calculation')

    ldf = study_case_obj.SearchObject('*.ComLdf')                                  # Get load flow calculation object




    o_IntPrjFolder_netdat = app.GetProjectFolder('netdat')
    o_ElmNet = o_IntPrjFolder_netdat.SearchObject('*.ElmNet')                   # Get network

    app.Hide()                                                                  # Hide GUI of powerfactory

    return app, study_case_obj, ldf, o_ElmNet

def set_load_flow_settings(ldf_com_obj, load_scaling, generation_scaling):

    ldf_com_obj.SetAttribute('scLoadFac', load_scaling)          # set load scaling factor
    ldf_com_obj.SetAttribute('scGenFac', generation_scaling)     # set generation scaling factor

    ldf_com_obj.SetAttribute('iopt_net', 0)  # AC Load Flow, balanced, positive sequence
    ldf_com_obj.SetAttribute('iopt_at', 0)  # Automatic tap adjustment of transformers
    ldf_com_obj.SetAttribute('iopt_pq', 0)  # Consider Voltage Dependency of Loads

    return