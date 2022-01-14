import helpers

def processar():

    jsonResponse = helpers.buscaVagas()
    if (jsonResponse == False):
        return False
    
    if (helpers.montarPlanilha(jsonResponse) == False):
        return False
    
    if (helpers.enviarEmail() == False):
        return False


processar()