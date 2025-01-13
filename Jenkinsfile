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
                    // Set up a virtual environment and activate it
                    sh 'python3 -m venv venv'  // Create virtual environment
                    sh '. venv/bin/activate'    // Activate virtual environment
                    sh 'pip install -r requirements.txt'  // Install dependencies
                }
            }
        }

        stage('Run Python Script') {
            steps {
                script {
                    sh '. venv/bin/activate && python main.py'  // Run the Python script
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
