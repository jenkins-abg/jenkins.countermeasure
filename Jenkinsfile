/****************************************************************/
/*    jenkins-abg = JENKINS FILE                                */
/*                                                              */
/*    (c) jenkins-abg - https://github.com/jenkins-abg          */      
/*                                                              */
/*    Copyright (c) 2020, Paul Ryan Dedumo                      */
/*    All rights reserved.                                      */           
/*                                                              */
/*    This file is coded for specific purposes only.            */
/*    Modifying this code may result to tool malfunctioning.    */
/*                                                              */
/****************************************************************/
library 'jenkinsci-unstashParam-library' 

def _myArrayName = [] 
def index = 0
def relPath = 'C:/work/Jenkins/automation-jenkins'


node("${slaveName}") { 
    def file_in_workspace = unstashParam "environmentName" 
    def mySlave =  "${slaveName}"
    def myParentFolder =  ""
    def myLocalIP = ""
    echo "copying to workspace... ${file_in_workspace}" 
}


//Checking for settings if user selected a file

if("${environmentName}" == ""){
    echo "No environment selected..."
    currentBuild.result = 'SUCCESS'
    return
}else{
    echo "executing pipeline..."
}

def environmentSplitter() {
    def envName = readFile("${environmentName}")
    return envName        
}

pipeline {
    
    environment {
        FILENAME = environmentSplitter()
        FILELOG = fileExists (relPath + '/Log.txt')
        NODE_NAME = "${slaveName}"
    }

    agent {
        node {
            label "${slaveName}"
            //unstashParam "environmentName" 
        }
    }
    /*stages {
        stage ("print slave") {
            steps {
                echo "${slaveName}"
            }
        }
    }*/
    
    stages{
        stage ('Reading CSV') {
            agent {
                node {
                    label "${slaveName}"
                }
            }
            steps {
                script {
                    def settings = "${env.FILENAME}"
                    settings.split('\n').each { line, count ->
                    def fields = line.split(',')
                        node {
                            echo fields[0] + ': ' + fields[1] ;
                            //store data in array
                            _myArrayName[index] = fields[0] + ': ' + fields[1] ;    
                        }
                        index++;
                    }
                    if("${env.FILELOG}" == 'true'){
                        echo "file exists"
                    }else{
                        targetFile = new File(relPath + '/Log.txt');
                        echo "no file exists..."
                        echo "creating file..."
                    }
                    writeFile file: (relPath + '/Log.txt'), text: ("""${_myArrayName[4]}\n${_myArrayName[0]}\nError_Number:\t0\nStatus:\t Starting\nSpider Version:\t\nSpider Log Text:\t\nError##_Fix:\t""")
                }
            }
        }
        stage('Clear Folder...'){   
            agent {
                node {
                    label "${slaveName}"
                }
            }
            steps {
                build job: '_jenkins_ClearDataInit', parameters:([
                        [$class: 'LabelParameterValue', name: 'node', label: "${slaveName}"]
                        ])
                //build job: '_jenkins_Build', quietPeriod: 1
            }
        }
        stage('Creating remote environment...'){
            agent {
                node {
                    label "${slaveName}"
                }
            }
            steps {
                build job: '_jenkins_CopyLocalToRemote', parameters:([
                        [$class: 'StringParameterValue', name: 'mySlave', value: "${slaveName}"],
                        [$class: 'StringParameterValue', name: 'myLocalIP', value: "${_myArrayName[9]}"],
                        [$class: 'StringParameterValue', name: 'myTestSheet', value: "${_myArrayName[3]}"],
                        [$class: 'StringParameterValue', name: 'myTestSheetParentFolder', value: "${_myArrayName[1]}"],
                        [$class: 'StringParameterValue', name: 'mySoftwareOneFolderUp', value: "${_myArrayName[5]}"],
                        [$class: 'StringParameterValue', name: 'mySoftwareParentFolder', value: "${_myArrayName[6]}"],
                        [$class: 'LabelParameterValue', name: 'node', label: "${slaveName}"]
                        ])
            }
        }
        stage('Initializing...'){
            agent {
                node {
                    label "${slaveName}"
                }
            }
            options {
                timeout(time: 5, unit: "MINUTES")
            }
            steps {
                    build job: '_jenkins_Main', parameters:([
                        [$class: 'LabelParameterValue', name: 'node', label: "${slaveName}"]
                        ])
                    //build job: '_jenkins_Build', quietPeriod: 1
            }
        }     
    }       
}

