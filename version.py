class Version:
    branch = None
    build = None
    cup = None
    
    def __init__(self, branch, build, cpu):
        self.branch = branch
        self.build = build
        self.cpu = cpu
