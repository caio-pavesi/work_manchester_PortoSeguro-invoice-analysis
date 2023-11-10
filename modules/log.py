from pathlib import Path as ph

class Log:
    ## Mutable vars
    output_file: ph = None
    
    ## Immutable vars
    def __init__(self):
        pass
    
    ## Creates a "log" in file.
    def log(self, information: str):
        print(information)
        file = self.output_file
        log = open(file, 'a')
        log.write(information)
        log.close()
        
        return log