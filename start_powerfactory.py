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