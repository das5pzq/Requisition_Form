import unittest
from app import app
from datetime import date

# Google Form URL (Change to your form URL)
GOOGLE_FORM_URL = "https://forms.gle/Z16J2t2GtqtP9wJN8"

# Email credentials
EMAIL_SENDER = "das5pzq@virginia.edu"
EMAIL_PASSWORD = "Landau0911!!"
EMAIL_RECEIVER = "das5pzq@virginia.edu"

class TestPurchaseOrderApp(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()
        
        # Sample valid form data matching your HTML structure
        self.valid_data = {
            'name': 'John Doe',
            'vendor': 'Test Vendor',
            'date': '2024-03-20',
            'telephone': '123-456-7890',
            'fax': '123-456-7899',  # Optional
            'ptao': 'PT123456',
            'purpose': 'Testing purposes',
            'catalog_number[]': ['CAT001', 'CAT002'],
            'item_name[]': ['Item 1', 'Item 2'],
            'item_quantity[]': ['1', '2'],
            'item_price[]': ['10.00', '20.00']
        }

    def test_home_page_loads(self):
        """Test if the form page loads correctly"""
        response = self.client.get('/')
        self.assertEqual(response.status_code, 200)
        # Check for key elements in the form
        self.assertIn(b'Purchase Order Form', response.data)
        self.assertIn(b'Name of Purchaser', response.data)
        self.assertIn(b'Vendor Name', response.data)

    def test_missing_required_fields(self):
        """Test submission with missing required fields"""
        required_fields = ['name', 'vendor', 'date', 'telephone', 'ptao', 'purpose']
        
        for field in required_fields:
            invalid_data = self.valid_data.copy()
            del invalid_data[field]
            response = self.client.post('/', data=invalid_data)
            self.assertEqual(response.status_code, 400, f"Should fail when {field} is missing")

    def test_invalid_item_data(self):
        """Test various invalid item data scenarios"""
        test_cases = [
            {
                'name': 'negative quantity',
                'data': {
                    **self.valid_data,
                    'item_quantity[]': ['-1', '2']
                }
            },
            {
                'name': 'negative price',
                'data': {
                    **self.valid_data,
                    'item_price[]': ['-10.00', '20.00']
                }
            },
            {
                'name': 'mismatched items',
                'data': {
                    **self.valid_data,
                    'item_name[]': ['Item 1'],  # Only one name but multiple quantities/prices
                    'item_quantity[]': ['1', '2'],
                    'item_price[]': ['10.00', '20.00']
                }
            }
        ]

        for test_case in test_cases:
            response = self.client.post('/', data=test_case['data'])
            self.assertEqual(
                response.status_code, 
                400, 
                f"Should fail with {test_case['name']}"
            )

    def test_valid_submission(self):
        """Test a complete valid form submission"""
        response = self.client.post('/', data=self.valid_data)
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Purchase order processed successfully!", response.data)

    @patch('yagmail.SMTP')
    def test_email_sending(self, mock_smtp):
        """Test email sending functionality"""
        mock_smtp.return_value.send = MagicMock()
        response = self.client.post('/', data=self.valid_data)
        self.assertEqual(response.status_code, 200)
        mock_smtp.return_value.send.assert_called_once()

    def test_fax_optional(self):
        """Test that fax field is optional"""
        data = self.valid_data.copy()
        del data['fax']
        response = self.client.post('/', data=data)
        self.assertEqual(response.status_code, 200)

    def test_multiple_items(self):
        """Test submission with multiple items"""
        data = self.valid_data.copy()
        data.update({
            'catalog_number[]': ['CAT001', 'CAT002', 'CAT003'],
            'item_name[]': ['Item 1', 'Item 2', 'Item 3'],
            'item_quantity[]': ['1', '2', '3'],
            'item_price[]': ['10.00', '20.00', '30.00']
        })
        response = self.client.post('/', data=data)
        self.assertEqual(response.status_code, 200)

if __name__ == '__main__':
    unittest.main() 