let currentSection = 1;
const totalSections = 7;
let assessmentData = {};

// Question mapping to categories
const questionCategories = {
    jobDemands: [1, 2, 3, 4, 5],
    control: [6, 7, 8, 9, 10],
    managerSupport: [11, 12, 13, 14, 15],
    peerSupport: [16, 17, 18, 19, 20],
    relationships: [21, 22, 23, 24, 25],
    roleClarity: [26, 27, 28, 29, 30],
    changeManagement: [31, 32, 33, 34, 35]
};

function startAssessment() {
    // Get optional info
    assessmentData.name = document.getElementById('assessor-name').value || 'Anonymous';
    assessmentData.department = document.getElementById('department').value || 'Not specified';
    assessmentData.date = new Date().toLocaleDateString();

    // Switch screens
    document.getElementById('welcome-screen').classList.remove('active');
    document.getElementById('assessment-screen').classList.add('active');

    updateProgress();
}

function updateProgress() {
    const progress = (currentSection / totalSections) * 100;
    document.getElementById('progress-fill').style.width = progress + '%';
    document.getElementById('current-section').textContent = currentSection;
    document.getElementById('total-sections').textContent = totalSections;
}

function nextSection() {
    // Validate current section
    if (!validateSection(currentSection)) {
        alert('Please answer all questions in this section before proceeding.');
        return;
    }

    // Hide current section
    document.querySelector(`.section[data-section="${currentSection}"]`).style.display = 'none';

    // Move to next section
    currentSection++;

    if (currentSection <= totalSections) {
        // Show next section
        document.querySelector(`.section[data-section="${currentSection}"]`).style.display = 'block';
        updateProgress();

        // Update button visibility
        document.getElementById('prev-btn').style.display = 'inline-block';

        if (currentSection === totalSections) {
            document.getElementById('next-btn').style.display = 'none';
            document.getElementById('submit-btn').style.display = 'inline-block';
        }

        // Scroll to top
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
}

function previousSection() {
    // Hide current section
    document.querySelector(`.section[data-section="${currentSection}"]`).style.display = 'none';

    // Move to previous section
    currentSection--;

    // Show previous section
    document.querySelector(`.section[data-section="${currentSection}"]`).style.display = 'block';
    updateProgress();

    // Update button visibility
    if (currentSection === 1) {
        document.getElementById('prev-btn').style.display = 'none';
    }

    document.getElementById('next-btn').style.display = 'inline-block';
    document.getElementById('submit-btn').style.display = 'none';

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function validateSection(section) {
    const sectionElement = document.querySelector(`.section[data-section="${section}"]`);
    const questions = sectionElement.querySelectorAll('.question');

    for (let question of questions) {
        const radios = question.querySelectorAll('input[type="radio"]');
        const questionName = radios[0].name;
        const isAnswered = Array.from(radios).some(radio => radio.checked);

        if (!isAnswered) {
            return false;
        }
    }

    return true;
}

function submitAssessment() {
    // Validate final section
    if (!validateSection(currentSection)) {
        alert('Please answer all questions before viewing results.');
        return;
    }

    // Collect all answers
    const form = document.getElementById('assessment-form');
    const formData = new FormData(form);

    assessmentData.answers = {};
    for (let [key, value] of formData.entries()) {
        assessmentData.answers[key] = parseInt(value);
    }

    // Calculate results
    calculateResults();

    // Show results screen
    document.getElementById('assessment-screen').classList.remove('active');
    document.getElementById('results-screen').classList.add('active');

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function calculateResults() {
    const categories = {
        jobDemands: {
            name: 'Job Demands',
            questions: questionCategories.jobDemands,
            score: 0,
            maxScore: 25,
            description: ''
        },
        control: {
            name: 'Control & Autonomy',
            questions: questionCategories.control,
            score: 0,
            maxScore: 25,
            description: ''
        },
        managerSupport: {
            name: 'Manager Support',
            questions: questionCategories.managerSupport,
            score: 0,
            maxScore: 25,
            description: ''
        },
        peerSupport: {
            name: 'Peer Support',
            questions: questionCategories.peerSupport,
            score: 0,
            maxScore: 25,
            description: ''
        },
        relationships: {
            name: 'Relationships',
            questions: questionCategories.relationships,
            score: 0,
            maxScore: 25,
            description: ''
        },
        roleClarity: {
            name: 'Role Clarity',
            questions: questionCategories.roleClarity,
            score: 0,
            maxScore: 25,
            description: ''
        },
        changeManagement: {
            name: 'Change Management',
            questions: questionCategories.changeManagement,
            score: 0,
            maxScore: 25,
            description: ''
        }
    };

    // Calculate scores for each category
    for (let categoryKey in categories) {
        const category = categories[categoryKey];
        let totalScore = 0;

        for (let questionNum of category.questions) {
            const questionKey = 'q' + questionNum;
            totalScore += assessmentData.answers[questionKey] || 0;
        }

        category.score = totalScore;
        category.percentage = (totalScore / category.maxScore) * 100;
        category.rating = getRating(category.percentage);
        category.description = getCategoryDescription(categoryKey, category.rating);
    }

    // Calculate overall score
    let totalScore = 0;
    let maxTotalScore = 0;

    for (let categoryKey in categories) {
        totalScore += categories[categoryKey].score;
        maxTotalScore += categories[categoryKey].maxScore;
    }

    const overallPercentage = (totalScore / maxTotalScore) * 100;
    const overallRating = getRating(overallPercentage);

    // Display results
    displayResults(categories, overallPercentage, overallRating);
}

function getRating(percentage) {
    if (percentage >= 80) return 'excellent';
    if (percentage >= 60) return 'good';
    if (percentage >= 40) return 'moderate';
    if (percentage >= 20) return 'poor';
    return 'critical';
}

function getRatingLabel(rating) {
    const labels = {
        excellent: 'Excellent',
        good: 'Good',
        moderate: 'Needs Attention',
        poor: 'Poor',
        critical: 'Critical'
    };
    return labels[rating];
}

function getCategoryDescription(categoryKey, rating) {
    const descriptions = {
        jobDemands: {
            excellent: 'Your workload is manageable and well-balanced. You have sufficient time to complete tasks to a high standard.',
            good: 'Your workload is generally manageable, though there may be occasional periods of high pressure.',
            moderate: 'You are experiencing some difficulty managing your workload. Consider discussing priorities with your manager.',
            poor: 'Your workload is creating significant stress. Immediate action is needed to address work demands.',
            critical: 'Your workload is critically high and poses serious risks to your health and wellbeing. Urgent intervention required.'
        },
        control: {
            excellent: 'You have excellent autonomy and control over your work, which supports your wellbeing.',
            good: 'You have good control over how and when you do your work.',
            moderate: 'Your level of control over your work could be improved. Discuss ways to increase autonomy with your manager.',
            poor: 'You have limited control over your work, which may be causing stress and frustration.',
            critical: 'You have very little control over your work. This lack of autonomy is a significant psychosocial hazard.'
        },
        managerSupport: {
            excellent: 'You receive excellent support from your manager, which is a key protective factor for wellbeing.',
            good: 'Your manager provides good support and guidance.',
            moderate: 'Manager support could be improved. Consider requesting regular check-ins and feedback.',
            poor: 'You are not receiving adequate support from your manager. This needs to be addressed.',
            critical: 'The lack of manager support is a critical issue affecting your wellbeing and performance.'
        },
        peerSupport: {
            excellent: 'You have strong collegial support, which is excellent for team morale and wellbeing.',
            good: 'You receive good support from your colleagues.',
            moderate: 'Peer support could be strengthened. Consider team-building activities or buddy systems.',
            poor: 'You are not receiving adequate support from colleagues. Team dynamics need attention.',
            critical: 'The lack of peer support is creating significant isolation and stress.'
        },
        relationships: {
            excellent: 'Workplace relationships are positive and respectful, creating a healthy work environment.',
            good: 'Workplace relationships are generally positive with minor issues.',
            moderate: 'There are some concerning relationship issues that should be addressed proactively.',
            poor: 'Workplace relationships are strained and may involve conflict or disrespect.',
            critical: 'There are serious relationship problems including potential bullying or harassment. Immediate action required.'
        },
        roleClarity: {
            excellent: 'You have excellent clarity about your role, responsibilities, and expectations.',
            good: 'Your role is generally clear with minor areas for clarification.',
            moderate: 'Some aspects of your role are unclear. Request clarification from your manager.',
            poor: 'You lack clarity about your role and responsibilities, causing confusion and stress.',
            critical: 'Severe role ambiguity and conflicting demands are creating significant problems.'
        },
        changeManagement: {
            excellent: 'Organizational changes are managed very well with good communication and consultation.',
            good: 'Changes are generally managed well with adequate communication.',
            moderate: 'Change management could be improved with better communication and consultation.',
            poor: 'Changes are poorly communicated and managed, creating uncertainty and stress.',
            critical: 'The way changes are managed is creating severe stress and undermining trust in leadership.'
        }
    };

    return descriptions[categoryKey][rating];
}

function displayResults(categories, overallPercentage, overallRating) {
    // Display assessment info
    const infoHTML = `
        <p><strong>Assessed by:</strong> ${assessmentData.name}</p>
        <p><strong>Department:</strong> ${assessmentData.department}</p>
        <p><strong>Date:</strong> ${assessmentData.date}</p>
    `;
    document.getElementById('assessment-info').innerHTML = infoHTML;

    // Display overall score
    document.getElementById('overall-score').textContent = Math.round(overallPercentage) + '%';
    document.getElementById('overall-rating').textContent = getRatingLabel(overallRating);

    // Create visual chart
    let chartHTML = '<h3>Category Breakdown</h3>';
    for (let categoryKey in categories) {
        const category = categories[categoryKey];
        chartHTML += `
            <div class="chart-bar">
                <div class="chart-label">
                    <span>${category.name}</span>
                    <span>${Math.round(category.percentage)}%</span>
                </div>
                <div class="chart-bar-container">
                    <div class="chart-bar-fill ${category.rating}" style="width: ${category.percentage}%">
                    </div>
                </div>
            </div>
        `;
    }
    document.getElementById('results-chart').innerHTML = chartHTML;

    // Display category details
    let categoryHTML = '<h3>Detailed Category Analysis</h3>';
    for (let categoryKey in categories) {
        const category = categories[categoryKey];
        categoryHTML += `
            <div class="category-card">
                <div class="category-header">
                    <div class="category-title">${category.name}</div>
                    <div class="category-score">
                        <span class="score-value">${Math.round(category.percentage)}%</span>
                        <span class="risk-badge ${category.rating}">${getRatingLabel(category.rating)}</span>
                    </div>
                </div>
                <p class="category-description">${category.description}</p>
            </div>
        `;
    }
    document.getElementById('category-results').innerHTML = categoryHTML;

    // Generate recommendations
    generateRecommendations(categories);
}

function generateRecommendations(categories) {
    let recommendations = [];

    // Identify areas needing attention
    const priorityAreas = [];
    for (let categoryKey in categories) {
        const category = categories[categoryKey];
        if (category.rating === 'critical' || category.rating === 'poor') {
            priorityAreas.push({ key: categoryKey, ...category });
        }
    }

    // Sort by score (lowest first)
    priorityAreas.sort((a, b) => a.percentage - b.percentage);

    if (priorityAreas.length === 0) {
        recommendations.push({
            title: 'Overall Positive Assessment',
            items: [
                'Your workplace psychosocial environment is in good shape',
                'Continue to monitor these factors and maintain open communication',
                'Consider sharing best practices with other teams',
                'Regular reassessment every 6-12 months is recommended'
            ]
        });
    } else {
        recommendations.push({
            title: 'Priority Actions Required',
            items: [
                `${priorityAreas.length} area(s) require immediate attention`,
                'Schedule a meeting with your manager or HR to discuss concerns',
                'Document specific incidents or patterns that contribute to low scores',
                'Request a workplace risk assessment if issues persist'
            ]
        });

        // Specific recommendations for each priority area
        const specificRecommendations = {
            jobDemands: {
                title: 'Managing Job Demands',
                items: [
                    'Review and prioritize tasks with your manager',
                    'Discuss realistic deadlines and workload distribution',
                    'Identify tasks that could be delegated or streamlined',
                    'Ensure you take regular breaks throughout the day',
                    'Consider time management training or tools'
                ]
            },
            control: {
                title: 'Increasing Control & Autonomy',
                items: [
                    'Request more involvement in decisions affecting your work',
                    'Discuss flexible working arrangements with your manager',
                    'Seek opportunities to use your skills and expertise',
                    'Propose alternative work methods if current processes are inefficient',
                    'Participate in planning and goal-setting discussions'
                ]
            },
            managerSupport: {
                title: 'Improving Manager Support',
                items: [
                    'Request regular one-on-one meetings with your manager',
                    'Clearly communicate when you need support or guidance',
                    'Provide feedback on what type of support would be most helpful',
                    'If issues persist, consider speaking with HR or a senior manager',
                    'Explore mentoring programs within your organization'
                ]
            },
            peerSupport: {
                title: 'Building Peer Support',
                items: [
                    'Participate in team activities and social events',
                    'Offer help and support to colleagues when you can',
                    'Join or establish a peer support network or buddy system',
                    'Suggest team-building activities to your manager',
                    'Be open about challenges and encourage others to do the same'
                ]
            },
            relationships: {
                title: 'Addressing Relationship Issues',
                items: [
                    'Document any incidents of bullying, harassment, or conflict',
                    'Speak with your manager or HR about relationship concerns',
                    'Review your organization\'s policies on respectful workplace behavior',
                    'Consider mediation or conflict resolution services',
                    'If you feel unsafe, report to appropriate authorities immediately'
                ]
            },
            roleClarity: {
                title: 'Clarifying Your Role',
                items: [
                    'Request a formal review of your position description',
                    'Ask for clear, written expectations and priorities',
                    'Discuss how your role contributes to organizational goals',
                    'Seek clarification when you receive conflicting instructions',
                    'Document agreed responsibilities and refer to them regularly'
                ]
            },
            changeManagement: {
                title: 'Improving Change Management',
                items: [
                    'Request more information about planned changes',
                    'Ask to participate in consultation processes',
                    'Provide feedback through official channels',
                    'Join or establish employee consultation committees',
                    'Document how changes are affecting your work and wellbeing'
                ]
            }
        };

        // Add specific recommendations for top 3 priority areas
        for (let i = 0; i < Math.min(3, priorityAreas.length); i++) {
            const area = priorityAreas[i];
            if (specificRecommendations[area.key]) {
                recommendations.push(specificRecommendations[area.key]);
            }
        }
    }

    // Add general wellbeing recommendations
    recommendations.push({
        title: 'General Wellbeing Support',
        items: [
            'Maintain work-life balance and disconnect from work outside hours',
            'Use Employee Assistance Programs (EAP) if available',
            'Practice stress management techniques like mindfulness or exercise',
            'Connect with colleagues and build supportive relationships',
            'Speak up about workplace issues before they escalate',
            'Familiarize yourself with workplace health and safety resources'
        ]
    });

    // Display recommendations
    let recommendationsHTML = '';
    for (let rec of recommendations) {
        recommendationsHTML += `
            <div class="recommendation-card">
                <h4>${rec.title}</h4>
                <ul>
                    ${rec.items.map(item => `<li>${item}</li>`).join('')}
                </ul>
            </div>
        `;
    }
    document.getElementById('recommendations-list').innerHTML = recommendationsHTML;
}

function printReport() {
    window.print();
}

function restartAssessment() {
    // Reset everything
    currentSection = 1;
    assessmentData = {};

    // Reset form
    document.getElementById('assessment-form').reset();
    document.getElementById('assessor-name').value = '';
    document.getElementById('department').value = '';

    // Hide all sections except first
    document.querySelectorAll('.section').forEach((section, index) => {
        section.style.display = index === 0 ? 'block' : 'none';
    });

    // Reset buttons
    document.getElementById('prev-btn').style.display = 'none';
    document.getElementById('next-btn').style.display = 'inline-block';
    document.getElementById('submit-btn').style.display = 'none';

    // Show welcome screen
    document.getElementById('results-screen').classList.remove('active');
    document.getElementById('welcome-screen').classList.add('active');

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
