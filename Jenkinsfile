pipeline {
    agent any

    environment {
        // Set Python version or path if necessary
        PYTHON = 'python3'  // or 'python' based on your environment
    }

    stages {
        stage('Checkout') {
            steps {
                // Pull the latest code from the GitHub repository
                git branch: 'main', url: 'https://github.com/Chaitrakotagiri/GUInterface.git'
            }
        }

        stage('Setup Python Environment') {
            steps {
                // Install Python dependencies if necessary (e.g., from requirements.txt)
                script {
                    sh 'pip install -r requirements.txt'  // Adjust if using a different file or package manager
                }
            }
        }

        stage('Run Python Script') {
            steps {
                // Run the main.py script
                script {
                    sh 'python main.py'
                }
            }
        }
    }

    post {
        always {
            // Clean up or perform other actions like sending notifications
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
