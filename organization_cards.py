"""Responsive, static department hierarchy; height follows its actual content."""
from html import escape


def _card(department, *, root=False):
    leaders = ', '.join(department['leaders']) or 'Responsable à désigner'
    members = ''.join(f'<li>{escape(str(name))}</li>' for name in department['members'])
    if not members:
        members = '<li class="org-muted">Aucun membre affecté</li>'
    return (
        f'<article class="org-card{" org-root" if root else ""}">'
        f'<h3>{escape(str(department["name"]))}</h3>'
        '<div class="org-label">Responsable</div>'
        f'<p class="org-leader">{escape(leaders)}</p>'
        '<div class="org-label">Équipe</div>'
        f'<ul>{members}</ul></article>'
    )


def organization_html(model):
    general = next((d for d in model['departments'] if d['id'] == model['general_id']), None)
    general = general or {'name': 'Direction Générale · à configurer', 'leaders': [], 'members': []}
    children = [d for d in model['departments'] if d['id'] != model['general_id']]
    css = '''<style>
    .org-responsive{width:100%;box-sizing:border-box;padding:8px 0 0;color:inherit;
      container-type:inline-size}
    .org-responsive *{box-sizing:border-box}
    .org-card{min-width:0;border:1px solid #455468;border-radius:12px;
      padding:18px;background:rgba(100,130,170,.055);position:relative}
    .org-card h3{font-size:18px;line-height:1.35;font-weight:700;margin:0 0 16px;
      padding-bottom:12px;border-bottom:1px solid #455468;overflow-wrap:anywhere}
    .org-label{text-transform:uppercase;letter-spacing:.08em;font-size:10px;
      color:#94a3b8;margin:12px 0 6px;line-height:1.4}
    .org-leader{font-size:14px;color:#65b2d0;line-height:1.5;margin:0;
      overflow-wrap:anywhere;font-weight:600}
    .org-card ul{list-style:none;padding:0;margin:0;font-size:14px;line-height:1.6}
    .org-card li{overflow-wrap:anywhere}
    .org-muted{color:#94a3b8;font-size:13px}
    .org-root{width:min(100%,420px);margin:0 auto;border-color:#65b2d0;
      background:rgba(80,160,190,.08)}
    .org-root h3{font-size:22px}
    .org-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));
      gap:16px;margin-top:24px}
    @container (max-width:780px){.org-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
    @container (max-width:480px){.org-grid{grid-template-columns:minmax(0,1fr)}}
    </style>'''
    bottom = '<div class="org-grid">' + ''.join(_card(d) for d in children) + '</div>' if children else ''
    return css + '<section class="org-responsive" aria-label="Organigramme des départements">' + _card(general, root=True) + bottom + '</section>'
