"""
test_layouts.py — Génère test_output.pptx avec une slide par layout.
Usage : python test_layouts.py
"""
from pptx import Presentation
from pptx.util import Inches
from layouts import LAYOUT_REGISTRY

PALETTE = {
    'primary':   '1A3A6B',
    'secondary': '2E6DA4',
    'accent':    'F0A500',
    'light':     'E8EEF7',
    'text':      '0A1525',
    'font':      'Calibri',
}

TEST_CONTENT = {
    'cover_dark': {
        'title':    'Stratégie Énergétique de la France',
        'subtitle': 'Analyse & perspectives 2024–2030',
        'footer':   'Visual Cortex · 2024',
    },
    'cover_split': {
        'title':    'Transition Énergétique',
        'subtitle': 'Les défis et opportunités pour les entreprises françaises',
        'footer':   'Visual Cortex · 2024',
    },
    'section': {
        'number': '01',
        'title':  'Contexte Énergétique Européen',
    },
    'kpi_grid': {
        'title': 'Chiffres Clés du Secteur',
        'kpis': [
            {'value': '70 %',   'label': 'Part du nucléaire',    'sublabel': 'Production électrique 2023'},
            {'value': '33 %',   'label': 'Objectif ENR 2030',    'sublabel': 'Programmation pluriannuelle'},
            {'value': '−38 %',  'label': 'Réduction CO₂',        'sublabel': 'vs 1990, objectif 2030'},
            {'value': '€ 40Md', 'label': 'Investissements',      'sublabel': 'Prévus sur 5 ans'},
            {'value': '6 EPR',  'label': 'Nouveaux réacteurs',   'sublabel': 'Plan nucléaire 2035'},
            {'value': '100 GW', 'label': 'Solaire 2050',         'sublabel': 'Capacité installée cible'},
        ],
        'footer': 'Visual Cortex · 2024',
    },
    'kpi_row': {
        'title': 'Indicateurs de Performance',
        'kpis': [
            {'value': '548 TWh', 'label': 'Production totale', 'sublabel': '2023'},
            {'value': '92 %',    'label': 'Disponibilité',     'sublabel': 'Parc nucléaire'},
            {'value': '−12 %',   'label': 'Consommation',      'sublabel': 'vs 2019'},
        ],
        'footer': 'Visual Cortex · 2024',
    },
    'timeline_h': {
        'title': "Programmation Pluriannuelle de l'Énergie",
        'steps': [
            {'date': '2019', 'title': 'PPE publiée',         'body': 'Lancement du plan national'},
            {'date': '2022', 'title': 'Choc énergétique',    'body': 'Crise gaz & électricité'},
            {'date': '2023', 'title': 'Relance nucléaire',   'body': 'Annonce 6 nouveaux EPR2'},
            {'date': '2025', 'title': 'PPE révisée',         'body': 'Nouveaux objectifs ENR'},
            {'date': '2030', 'title': 'Bilan CO₂ −38 %',    'body': 'Objectif européen'},
        ],
        'footer': 'Visual Cortex · 2024',
    },
    'two_col': {
        'title': 'Nucléaire vs Renouvelables',
        'col_a': {
            'title': 'Nucléaire',
            'items': [
                'Production pilotable 24/7',
                'Faibles émissions de CO₂',
                'Coût à long terme compétitif',
                'Délais de construction longs',
                'Gestion des déchets complexe',
            ],
        },
        'col_b': {
            'title': 'Renouvelables',
            'items': [
                'Déploiement rapide',
                'Coûts en forte baisse',
                'Production intermittente',
                'Besoin de stockage massif',
                'Impact paysager débattu',
            ],
        },
        'footer': 'Visual Cortex · 2024',
    },
    'quote_dark': {
        'quote':  'La France a une carte maîtresse à jouer dans la transition énergétique européenne.',
        'author': "Agnès Pannier-Runacher, Ministre de l'Énergie",
        'footer': 'Visual Cortex · 2024',
    },
    'list_numbered': {
        'title': 'Les 4 Leviers de la Transition',
        'items': [
            {'title': 'Sobriété énergétique',       'body': 'Réduire la consommation de 10 % via mesures comportementales et réglementaires.'},
            {'title': 'Efficacité des bâtiments',   'body': 'Rénover 700 000 logements par an, priorité aux passoires thermiques.'},
            {'title': 'Électrification des usages', 'body': 'Véhicules électriques, pompes à chaleur, industrie bas-carbone.'},
            {'title': 'Production décarbonée',      'body': "Mix nucléaire + ENR pour atteindre 100 % d'électricité bas-carbone."},
        ],
        'footer': 'Visual Cortex · 2024',
    },
    'list_cards': {
        'title': 'Axes Stratégiques',
        'cards': [
            {'title': "Sécurité d'approvisionnement", 'body': 'Diversifier les sources, réduire la dépendance aux importations.'},
            {'title': 'Compétitivité industrielle',    'body': 'Maintenir des prix compétitifs pour les industries énergivores.'},
            {'title': 'Neutralité carbone 2050',       'body': "Atteindre le net-zéro conformément à l'accord de Paris."},
            {'title': 'Innovation technologique',      'body': "Investir dans l'hydrogène vert, le stockage et les SMR."},
        ],
        'footer': 'Visual Cortex · 2024',
    },
    'image_split': {
        'title':  "L'Électricité au Cœur de la Transition",
        'points': [
            "Part de l'électricité : 25 % → 55 % de la consommation finale en 2050",
            'Doublement du réseau de distribution nécessaire d\'ici 2035',
            'Investissements RTE : 100 Md€ sur 15 ans',
            'Création de 300 000 emplois dans la filière électrique',
            'Enjeu de souveraineté industrielle et technologique',
        ],
        'footer': 'Visual Cortex · 2024',
    },
    'full_text': {
        'title': 'Analyse Stratégique',
        'paragraphs': [
            "La France dispose d'un avantage structurel avec son parc nucléaire historique, permettant une production bas-carbone à grande échelle. La relance de ce secteur répond aux enjeux climatiques et de souveraineté.",
            'Le développement des renouvelables s\'accélère sous l\'effet de la baisse des coûts et des obligations européennes. Le mix électrique devrait atteindre 100 % d\'énergie décarbonée dès 2035.',
            'Les défis restent majeurs : financement, compétences, adaptation des réseaux et acceptabilité sociale. La cohérence de la politique énergétique dans la durée est la condition du succès.',
        ],
        'footer': 'Visual Cortex · 2024',
    },
    'stat_hero': {
        'value':   '2 050',
        'label':   'Objectif neutralité carbone',
        'context': 'La France vise la neutralité carbone en 2050 — Stratégie Nationale Bas-Carbone.',
        'footer':  'Visual Cortex · 2024',
    },
    'closing_dark': {
        'title':    'Merci de votre attention',
        'subtitle': "Sources : RTE, ADEME, Ministère de l'Énergie — 2024",
    },
    'closing_split': {
        'title':    'Passons à l\'action',
        'subtitle': 'Contactez notre équipe pour un accompagnement personnalisé sur votre stratégie énergétique.',
    },
}


def main():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    ok, ko = 0, 0
    for name, fn in LAYOUT_REGISTRY.items():
        content = TEST_CONTENT.get(name, {'title': name})
        try:
            fn(prs, content, PALETTE)
            print(f'  OK  {name}')
            ok += 1
        except Exception as e:
            print(f'  ERR {name}: {e}')
            ko += 1

    out = 'test_output.pptx'
    prs.save(out)
    print(f'\n{ok} OK / {ko} erreurs — sauvegardé dans {out}')


if __name__ == '__main__':
    main()
