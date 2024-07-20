import requests
from dataclasses import dataclass

class ExplanationServiceAPI:
    def __init__(self, service_endpoint):
        self.endpoint = service_endpoint

    def submit_document(self, document_path):
        with open(document_path, 'rb') as document:
            response = requests.post(f"{self.endpoint}/upload", files={'document': document})
        
        response.raise_for_status()
        return response.json()['document_id']

    def check_progress(self, document_id):
        response = requests.get(f"{self.endpoint}/status", params={'id': document_id})
        response.raise_for_status()
        return response.json()

@dataclass
class ProcessingStatus:
    state: str
    original_name: str
    submission_time: str
    result: str

    @property
    def is_complete(self):
        return self.state == 'complete'

    @classmethod
    def from_dict(cls, data):
        return cls(
            state=data['state'],
            original_name=data['original_name'],
            submission_time=data['submission_time'],
            result=data['result']
        )