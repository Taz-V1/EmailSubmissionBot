import unittest
from unittest.mock import Mock, patch
from main import LinkedInApp, get_linkedin_profile_info  # Replace with actual module name

class TestLinkedInScraper(unittest.TestCase):
    def setUp(self):
        self.mock_driver = Mock()
        self.app = LinkedInApp()
        self.app.sheet_service = Mock()

    def simulate_profile(self, html_content):
        """Mock browser content for testing"""
        self.mock_driver.page_source = html_content
        self.mock_driver.find_element.return_value.text = "Mocked Element Text"
        self.mock_driver.current_url = "https://www.linkedin.com/in/test-profile"

    def test_internship_detection(self):
        # Simulate profile with internship position
        test_html = """
        <div data-view-name="profile-component-entity">
            <h1>John Doe</h1>
            <span class="experience-item">Software Engineering Intern</span>
            <span class="education-item">State University</span>
        </div>
        """
        self.simulate_profile(test_html)
        
        with patch('hunterio_api.HunterIOAPI.find_email') as mock_find_email:
            mock_find_email.return_value = {'email': 'test@company.com', 'confidence': 95}
            
            result = get_linkedin_profile_info(self.app, self.mock_driver, "mocked_url")
            
        self.assertTrue(result['intern'])
        self.assertEqual(result['position'], 'Software Engineering Intern')
        self.assertFalse(result['thai'])

    def test_thai_education_detection(self):
        # Simulate profile with Thai education
        test_html = """
        <div data-view-name="profile-component-entity">
            <h1>Jane Smith</h1>
            <span class="experience-item">Senior Developer</span>
            <span class="education-item">Chulalongkorn University</span>
        </div>
        """
        self.simulate_profile(test_html)
        
        with patch('hunterio_api.HunterIOAPI.find_email') as mock_find_email:
            mock_find_email.return_value = {'email': 'test@company.com', 'confidence': 95}
            
            result = get_linkedin_profile_info(self.app, self.mock_driver, "mocked_url")
            
        self.assertTrue(result['thai'])
        self.assertFalse(result['intern'])

    def test_combined_case(self):
        # Simulate profile with both internship and Thai background
        test_html = """
        <div data-view-name="profile-component-entity">
            <h1>สมชาย ใจดี</h1>
            <span class="experience-item">Data Science Apprentice</span>
            <span class="education-item">International School Bangkok</span>
            <span class="language-item">Thai (Native)</span>
        </div>
        """
        self.simulate_profile(test_html)
        
        with patch('hunterio_api.HunterIOAPI.find_email') as mock_find_email:
            mock_find_email.return_value = {'email': 'test@company.com', 'confidence': 95}
            
            result = get_linkedin_profile_info(self.app, self.mock_driver, "mocked_url")
            
        self.assertTrue(result['intern'])
        self.assertTrue(result['thai'])
        self.assertTrue(result['highschool'])

    def test_negative_case(self):
        # Simulate profile with no flags
        test_html = """
        <div data-view-name="profile-component-entity">
            <h1>John Smith</h1>
            <span class="experience-item">Senior Engineer</span>
            <span class="education-item">MIT</span>
        </div>
        """
        self.simulate_profile(test_html)
        
        with patch('hunterio_api.HunterIOAPI.find_email') as mock_find_email:
            mock_find_email.return_value = {'email': 'test@company.com', 'confidence': 95}
            
            result = get_linkedin_profile_info(self.app, self.mock_driver, "mocked_url")
            
        self.assertFalse(result['intern'])
        self.assertFalse(result['thai'])
        self.assertFalse(result['college'])

if __name__ == '__main__':
    unittest.main()