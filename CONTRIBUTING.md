# Contributing to Set-OutOfOffice

First off, thank you for considering contributing to Set-OutOfOffice! It's people like you that make this tool better for the IT admin community.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How Can I Contribute?](#how-can-i-contribute)
- [Development Setup](#development-setup)
- [Coding Standards](#coding-standards)
- [Commit Guidelines](#commit-guidelines)
- [Pull Request Process](#pull-request-process)

## Code of Conduct

This project and everyone participating in it is governed by a simple principle: **Be kind and professional**. We expect all contributors to:

- Use welcoming and inclusive language
- Be respectful of differing viewpoints and experiences
- Gracefully accept constructive criticism
- Focus on what is best for the community
- Show empathy towards other community members

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates. When creating a bug report, include:

**Bug Report Template:**
```markdown
**Describe the bug**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Run script with '...'
2. Enter '....'
3. See error

**Expected behavior**
What you expected to happen.

**Screenshots/Error Messages**
If applicable, add screenshots or paste error messages.

**Environment:**
 - OS: [e.g., Windows 11]
 - PowerShell Version: [e.g., 7.4]
 - Module Versions: [Run `Get-Module Microsoft.Graph* -ListAvailable`]

**Additional context**
Any other context about the problem.
```

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion, include:

**Enhancement Template:**
```markdown
**Is your feature request related to a problem?**
A clear description of what the problem is.

**Describe the solution you'd like**
A clear description of what you want to happen.

**Describe alternatives you've considered**
Other solutions or features you've considered.

**Additional context**
Any other context, screenshots, or examples.
```

### Your First Code Contribution

Unsure where to begin? Look for issues tagged with:
- `good first issue` - Simple issues perfect for newcomers
- `help wanted` - Issues where we need community help

## Development Setup

1. **Fork and Clone**
   ```powershell
   git clone https://github.com/your-username/Set-OutOfOffice.git
   cd Set-OutOfOffice
   ```

2. **Install Required Modules**
   ```powershell
   Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
   ```

3. **Create a Test Environment**
   - Use a test Microsoft 365 tenant if possible
   - Create test mailboxes for safe testing
   - Never test on production mailboxes initially

4. **Create a Feature Branch**
   ```powershell
   git checkout -b feature/YourFeatureName
   ```

## Coding Standards

### PowerShell Best Practices

1. **Use Approved Verbs**
   ```powershell
   # Good
   Get-UserData
   Set-Configuration
   
   # Bad
   Fetch-UserData
   Change-Configuration
   ```

2. **Follow PascalCase for Functions**
   ```powershell
   function Get-MailboxSettings { }  # Good
   function get_mailbox_settings { }  # Bad
   ```

3. **Use Verbose Parameter Names**
   ```powershell
   # Good
   function Set-MailboxOOO {
       param(
           [string]$Mailbox,
           [DateTime]$StartTime
       )
   }
   
   # Avoid
   function Set-MailboxOOO {
       param($mb, $st)
   }
   ```

4. **Add Comment-Based Help**
   ```powershell
   <#
   .SYNOPSIS
       Brief description
   .DESCRIPTION
       Detailed description
   .PARAMETER ParameterName
       Parameter description
   .EXAMPLE
       Example usage
   #>
   ```

5. **Error Handling**
   ```powershell
   # Always use try/catch for external API calls
   try {
       # API call
   }
   catch {
       Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
   }
   ```

6. **Write Meaningful Comments**
   ```powershell
   # Good: Explains WHY
   # Convert to UTC because Graph API requires ISO 8601 format
   $startTimeISO = $StartTime.ToUniversalTime()
   
   # Bad: States the obvious
   # Convert to UTC
   $startTimeISO = $StartTime.ToUniversalTime()
   ```

### Code Formatting

- **Indentation:** 4 spaces (no tabs)
- **Line Length:** Aim for 120 characters maximum
- **Braces:** Opening brace on same line
  ```powershell
  if ($condition) {
      # code
  }
  ```

### Testing

Before submitting:
1. Test with a single mailbox
2. Test with multiple mailboxes
3. Test error conditions (invalid dates, wrong email format)
4. Test with different PowerShell versions if possible
5. Verify no credentials or sensitive data in output

## Commit Guidelines

### Commit Message Format

```
<type>: <subject>

<body>

<footer>
```

### Types
- **feat:** New feature
- **fix:** Bug fix
- **docs:** Documentation changes
- **style:** Code style changes (formatting, etc.)
- **refactor:** Code refactoring
- **test:** Adding or updating tests
- **chore:** Maintenance tasks

### Examples

```
feat: Add support for HTML formatted OOO messages

- Allow users to enter HTML tags in messages
- Preserve formatting when sending to Graph API
- Add validation for basic HTML syntax

Closes #42
```

```
fix: Correct date parsing for non-US locales

Previously, the script only accepted MM/DD/YYYY format.
Now correctly handles DD/MM/YYYY format which is more
common globally.

Fixes #15
```

## Pull Request Process

1. **Update Documentation**
   - Update README.md if you change functionality
   - Update code comments
   - Add examples if introducing new features

2. **Test Thoroughly**
   - Ensure all existing functionality still works
   - Test your new feature with various inputs
   - Include test results in PR description

3. **Create the Pull Request**
   ```markdown
   ## Description
   Brief description of changes
   
   ## Type of Change
   - [ ] Bug fix
   - [ ] New feature
   - [ ] Documentation update
   
   ## Testing Done
   - Tested with X mailboxes
   - Tested error handling
   - Tested on PowerShell 7.x
   
   ## Screenshots (if applicable)
   
   ## Checklist
   - [ ] My code follows the style guidelines
   - [ ] I have commented my code
   - [ ] I have updated documentation
   - [ ] My changes generate no new warnings
   - [ ] I have tested my changes
   ```

4. **Respond to Feedback**
   - Address reviewer comments promptly
   - Make requested changes
   - Update the PR with explanations

5. **Squash Commits (if requested)**
   ```powershell
   git rebase -i HEAD~n  # where n is number of commits
   ```

## Recognition

Contributors will be recognized in:
- README.md Contributors section
- GitHub Contributors page
- Release notes for significant contributions

## Questions?

Feel free to:
- Open an issue with the `question` label
- Start a discussion in the Discussions tab
- Reach out to maintainers

## License

By contributing, you agree that your contributions will be licensed under the same MIT License that covers this project.

---

Thank you for contributing to Set-OutOfOffice! 🎉
