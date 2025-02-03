import json

class LoteService:
    
    @staticmethod
    def loading_lotes():
        with open("lib/lotes.json","r") as file:
            data = json.load(file)
            LoteService.lotes_2025 = data["2025"]
            LoteService.lotes_2026 = data["2026"]
    
    @staticmethod
    def searchYearLote(lote):
        if lote in LoteService.lotes_2025:
            return 2025
        elif lote in LoteService.lotes_2026:
            return 2026
        else:
            return None

LoteService.loading_lotes()