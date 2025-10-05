# Contributing to ClearSend

Thank you for your interest in contributing to ClearSend! This document provides guidelines and instructions for contributing.

## 🌟 Ways to Contribute

- **Report Bugs**: Submit detailed bug reports with steps to reproduce
- **Suggest Features**: Propose new features or improvements
- **Write Code**: Fix bugs or implement new features
- **Improve Documentation**: Enhance README, add examples, fix typos
- **Test**: Help test new releases and report issues

## 🚀 Getting Started

### Prerequisites

- Node.js 14+ and npm
- Git
- Microsoft Outlook (Desktop or Web)
- Code editor (VS Code recommended)

### Setup Development Environment

1. **Fork the repository**
   ```bash
   # Click the "Fork" button on GitHub
   ```

2. **Clone your fork**
   ```bash
   git clone https://github.com/YOUR_USERNAME/ClearSend.git
   cd ClearSend
   ```

3. **Add upstream remote**
   ```bash
   git remote add upstream https://github.com/fhuerta01/ClearSend.git
   ```

4. **Install dependencies**
   ```bash
   npm install
   ```

5. **Start development server**
   ```bash
   npm run dev-server
   ```

6. **Sideload the add-in**
   ```bash
   npm start
   ```

## 📝 Development Guidelines

### Code Style

- Use ES6+ JavaScript features
- Follow existing code formatting
- Use meaningful variable and function names
- Add comments for complex logic
- Keep functions small and focused

### Code Organization

```
src/
├── taskpane/
│   ├── taskpane.js       # UI logic and event handlers
│   ├── processors.js     # Pure processing functions
│   └── clearsend.css     # Styles
└── commands/
    └── commands.js       # Quick action functions
```

### Best Practices

1. **Privacy First - Non-Negotiable**:
   - NEVER add network calls that transmit email data
   - All processing must remain 100% client-side
   - No external APIs, analytics, or tracking
   - Email addresses must never leave the user's device
2. **Error Handling**: Always handle errors gracefully
3. **User Feedback**: Provide clear feedback for all actions
4. **Performance**: Optimize for large recipient lists (1000+ recipients)
5. **Testing**: Test with various email formats and edge cases

### Coding Standards

```javascript
// Good: Descriptive names and clear logic
function validateEmailFormat(email) {
  if (!email || typeof email !== 'string') {
    return { isValid: false, message: 'Invalid email' };
  }
  // ...
}

// Bad: Unclear names and no validation
function check(e) {
  return e.includes('@');
}
```

## 🔧 Making Changes

### Workflow

1. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes**
   - Write clean, documented code
   - Follow existing patterns
   - Test thoroughly

3. **Test your changes**
   ```bash
   # Run linting
   npm run lint

   # Build for production
   npm run build

   # Test in Outlook
   npm start
   ```

4. **Commit your changes**
   ```bash
   git add .
   git commit -m "feat: add amazing feature"
   ```

### Commit Message Guidelines

Use conventional commits format:

- `feat:` New feature
- `fix:` Bug fix
- `docs:` Documentation changes
- `style:` Code style changes (formatting, etc.)
- `refactor:` Code refactoring
- `test:` Adding or updating tests
- `chore:` Maintenance tasks

Examples:
```
feat: add external recipient warning
fix: resolve duplicate detection bug
docs: update installation instructions
```

## 🐛 Reporting Bugs

### Before Submitting

1. Check existing issues to avoid duplicates
2. Test with the latest version
3. Gather detailed information

### Bug Report Template

```markdown
**Describe the bug**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '....'
3. See error

**Expected behavior**
What you expected to happen.

**Screenshots**
If applicable, add screenshots.

**Environment:**
- Outlook Version: [e.g. Outlook 365]
- Platform: [e.g. Desktop/Web]
- OS: [e.g. Windows 10]
- Browser (if web): [e.g. Chrome 120]

**Additional context**
Any other relevant information.
```

## 💡 Suggesting Features

### Feature Request Template

```markdown
**Is your feature request related to a problem?**
A clear description of the problem.

**Describe the solution you'd like**
What you want to happen.

**Describe alternatives you've considered**
Other solutions you've thought about.

**Additional context**
Mockups, examples, or other relevant information.
```

## 🔍 Pull Request Process

### Before Submitting

1. **Update your branch**
   ```bash
   git fetch upstream
   git rebase upstream/main
   ```

2. **Run all checks**
   ```bash
   npm run lint
   npm run build
   npm run validate
   ```

3. **Test thoroughly**
   - Test in Outlook Desktop
   - Test in Outlook Web
   - Test with various recipient lists
   - Check edge cases

### PR Guidelines

1. **Title**: Use clear, descriptive titles
   ```
   ✅ feat: add CSV export functionality
   ❌ Update code
   ```

2. **Description**: Include:
   - What changes were made
   - Why the changes were needed
   - How to test the changes
   - Screenshots/GIFs if UI changes

3. **Linking**: Reference related issues
   ```markdown
   Fixes #123
   Related to #456
   ```

4. **Size**: Keep PRs focused and reasonably sized
   - One feature/fix per PR
   - Split large changes into multiple PRs

### PR Template

```markdown
## Description
Brief description of changes.

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
How has this been tested?

## Checklist
- [ ] Code follows project style
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] No console.log statements
- [ ] Tested in Outlook Desktop
- [ ] Tested in Outlook Web
```

## 🧪 Testing

### Manual Testing Checklist

- [ ] Sort functionality with various recipient formats
- [ ] Duplicate detection across To/CC/BCC
- [ ] Email validation with edge cases
- [ ] Internal domain prioritization
- [ ] External recipient removal
- [ ] Undo functionality
- [ ] CSV export
- [ ] Settings persistence
- [ ] Error handling

### Test Cases

```javascript
// Example test cases to verify
const testRecipients = [
  'john@example.com',
  'John Doe <john@example.com>',  // Duplicate
  'invalid-email',                 // Invalid
  'alice@company.com',             // Internal
  'bob@external.com'               // External
];
```

## 📚 Documentation

### What to Document

- New features and how to use them
- API changes
- Configuration options
- Breaking changes
- Migration guides

### Documentation Style

- Clear and concise
- Include code examples
- Add screenshots for UI features
- Keep language simple and accessible

## 🤔 Questions?

- **GitHub Discussions**: Ask questions and discuss ideas
- **GitHub Issues**: Report bugs or request features
- **Email**: Contact the maintainers

## 🙏 Code of Conduct

Please read and follow our [Code of Conduct](CODE_OF_CONDUCT.md).

## ⚖️ License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to ClearSend! 🎉
