pipeline {
    agent any

    environment {
        PYTHON = 'python3'
    }

    stages {
        stage('Checkout') {
            steps {
                git url: 'https://github.com/Chaitrakotagiri/GUInterface.git', branch: 'main'
            }
        }

        stage('Setup Python Environment') {
            steps {
                script {
                    // Create virtual environment
                    sh 'python3 -m venv venv'
                    
                    // Manually install pip if it's missing
                    sh 'python3 -m ensurepip --upgrade'
                    
                    // Activate virtual environment and install requirements
                    sh '. venv/bin/activate && pip install -r requirements.txt'
                }
            }
        }

        stage('Run Python Script') {
            steps {
                script {
                    // Run the Python script in the virtual environment
                    sh '. venv/bin/activate && python main.py'
                }
            }
        }
    }

    post {
        always {
            echo 'Pipeline finished.'
        }
        success {
            echo 'The pipeline ran successfully.'
        }
        failure {
            echo 'The pipeline failed.'
        }
    }
}
