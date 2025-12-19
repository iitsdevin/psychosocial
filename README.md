# Workplace Psychosocial Risk Assessment Tool

A comprehensive web-based application for assessing psychosocial risks and hazards in workplace environments.

## Overview

This tool helps identify and evaluate psychosocial risks in workplaces by assessing seven key domains:

1. **Job Demands** - Workload, time pressure, and emotional demands
2. **Control & Autonomy** - Decision-making authority and task flexibility
3. **Manager Support** - Guidance, feedback, and assistance from supervisors
4. **Peer Support** - Collaboration and support from colleagues
5. **Relationships** - Workplace culture, respect, and conflict management
6. **Role Clarity** - Understanding of responsibilities and expectations
7. **Change Management** - How organizational changes are communicated and managed

## Features

- **35-question comprehensive assessment** based on established psychosocial risk frameworks
- **Interactive interface** with progress tracking
- **Detailed results** with visual charts and category breakdowns
- **Risk ratings** (Excellent, Good, Needs Attention, Poor, Critical)
- **Personalized recommendations** based on assessment results
- **Printable reports** for documentation and discussion
- **Mobile-responsive design** for accessibility on any device
- **Privacy-focused** - all data stays in your browser, nothing is transmitted

## How to Use

### Quick Start

1. Open `index.html` in any modern web browser
2. Enter your name and department (optional)
3. Click "Start Assessment"
4. Answer all 35 questions across 7 sections
5. View your detailed results and recommendations

### Using a Local Web Server

For best results, serve the application using a local web server:

```bash
# Using Python 3
python -m http.server 8000

# Using Python 2
python -m SimpleHTTPServer 8000

# Using Node.js (with http-server installed)
npx http-server

# Using PHP
php -S localhost:8000
```

Then open `http://localhost:8000` in your browser.

### Assessment Instructions

1. **Prepare**: Find a quiet space where you can answer honestly
2. **Timeframe**: Consider your experiences over the past 3 months
3. **Rating Scale**: Each question uses a 5-point scale (Always, Often, Sometimes, Seldom, Never)
4. **Honesty**: Be truthful - this helps identify real areas for improvement
5. **Completion**: Answer all questions in each section before proceeding

## Understanding Results

### Overall Score

Your overall psychosocial health score is calculated from all 35 questions:

- **80-100%** - Excellent: Workplace psychosocial environment is very healthy
- **60-79%** - Good: Generally positive with some areas for minor improvement
- **40-59%** - Needs Attention: Several areas require action
- **20-39%** - Poor: Significant psychosocial risks present
- **0-19%** - Critical: Urgent intervention required

### Category Scores

Each of the 7 domains is scored separately to identify specific strengths and areas of concern.

### Recommendations

The tool provides:
- Priority actions based on low-scoring areas
- Specific strategies for improvement
- General wellbeing support resources

## Privacy & Data

- All assessment data is processed locally in your browser
- No information is transmitted to external servers
- No data is stored permanently unless you choose to save/print the report
- Results can be printed for your records

## Use Cases

### For Employees
- Self-assessment of workplace wellbeing
- Identify areas of concern to discuss with managers
- Track changes over time
- Support for workplace health discussions

### For Managers
- Understand team psychosocial risks
- Identify areas for improvement
- Support employee wellbeing initiatives
- Inform workplace health and safety planning

### For Organizations
- Conduct workplace risk assessments
- Support compliance with occupational health requirements
- Identify training and support needs
- Benchmark across teams or departments

## Technical Details

### Requirements

- Modern web browser (Chrome, Firefox, Safari, Edge)
- JavaScript enabled
- No external dependencies or internet connection required

### Files

- `index.html` - Main application structure
- `styles.css` - Styling and responsive design
- `script.js` - Assessment logic and results calculation
- `README.md` - Documentation

### Browser Compatibility

- Chrome/Edge 90+
- Firefox 88+
- Safari 14+
- Mobile browsers (iOS Safari, Chrome Mobile)

## Best Practices

### For Individual Use

1. Complete the assessment when you're calm and can think clearly
2. Be honest - the tool only works if you answer truthfully
3. Retake every 3-6 months to track changes
4. Use results to have informed discussions with your manager
5. Seek professional support if you score consistently low

### For Organizational Use

1. Ensure confidentiality and voluntary participation
2. Communicate the purpose clearly to employees
3. Aggregate results to protect individual privacy
4. Act on findings - don't assess without planning to improve
5. Combine with other workplace health initiatives
6. Provide support resources for those who need them

## Limitations

- This is a screening tool, not a diagnostic instrument
- Results should inform discussion and action, not replace professional assessment
- Cultural and organizational factors may affect interpretation
- Individual results may not represent team-wide issues
- Should be part of a broader workplace health and safety program

## Legal & Ethical Considerations

- Results should be used constructively, not punitively
- Individual data should remain confidential
- Organizations should obtain consent before organizational-level assessments
- Follow local workplace health and safety regulations
- Provide support pathways for those identifying risks
- Consider privacy laws when storing or sharing data

## Support Resources

If the assessment reveals significant concerns:

- **Speak with your manager or HR department**
- **Use Employee Assistance Programs (EAP)** if available
- **Contact workplace health and safety representatives**
- **Seek professional counseling or support**
- **Review workplace policies** on bullying, harassment, and wellbeing

## Contributing

This is an open-source project. Suggestions for improvement are welcome:

- Report issues or bugs
- Suggest additional questions or categories
- Improve accessibility
- Add translations
- Enhance visualization

## License

This project is provided as-is for educational and organizational use. Modify and adapt as needed for your specific workplace context.

## Acknowledgments

Based on established psychosocial risk assessment frameworks including:
- UK Health and Safety Executive (HSE) Management Standards
- Workplace psychosocial risk assessment principles
- Occupational health and safety best practices

## Version

Version 1.0.0 - Initial Release

---

**Remember**: Assessing psychosocial risks is the first step. Taking action to address identified issues is what creates healthier workplaces.