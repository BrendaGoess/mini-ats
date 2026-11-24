import csv
from pathlib import Path
T = [
("Interview invitation","EN","CV passed screening","Interview invitation – [Role] at [Company]",
"Hi [Name],\n\nThank you for applying for the [Role] position at [Company]. We enjoyed reading your CV and would like to invite you to an interview.\n\nDate: [Date]\nTime: [Time] (Lisbon time)\nFormat: [Video / In person]\nLink or address: [Link]\n\nThe interview will take about 45 minutes. Please confirm if this time works for you, or suggest another slot.\n\nBest regards,\n[Your name]\n[Company]"),
("Interview invitation","PT-PT","CV passed screening","Convite para entrevista – [Função] na [Empresa]",
"Olá [Nome],\n\nObrigada pela sua candidatura à função de [Função] na [Empresa]. Gostámos do seu CV e gostaríamos de a/o convidar para uma entrevista.\n\nData: [Data]\nHora: [Hora]\nFormato: [Videochamada / Presencial]\nLink ou morada: [Link]\n\nA entrevista terá a duração de cerca de 45 minutos. Pode confirmar se este horário lhe convém ou sugerir outro?\n\nCom os melhores cumprimentos,\n[O seu nome]\n[Empresa]"),
("Kind rejection","EN","Not moving forward (after CV or interview)","Your application for [Role] at [Company]",
"Hi [Name],\n\nThank you for your interest in the [Role] position and for the time you gave us.\n\nAfter careful review, we have decided to move forward with other candidates whose experience is a closer match for this role.\n\nWe were glad to get to know you, and we would like to keep your CV for future opportunities. We wish you the best of luck in your search.\n\nKind regards,\n[Your name]\n[Company]"),
("Kind rejection","PT-PT","Not moving forward (after CV or interview)","A sua candidatura – [Função] na [Empresa]",
"Olá [Nome],\n\nObrigada pelo seu interesse na função de [Função] e pelo tempo que nos dedicou.\n\nApós uma análise cuidada, decidimos avançar com outros candidatos cujo perfil está mais alinhado com esta função.\n\nFoi um gosto conhecê-la/o e, se concordar, guardaremos o seu CV para futuras oportunidades. Desejamos-lhe muito sucesso.\n\nCom os melhores cumprimentos,\n[O seu nome]\n[Empresa]"),
("Job offer","EN","Final decision: hire","Job offer – [Role] at [Company]",
"Hi [Name],\n\nWe are happy to offer you the [Role] position at [Company]!\n\nMain terms:\n- Start date: [Date]\n- Contract: [Type]\n- Salary: [Amount] € gross per year\n- Work model: [Remote / Hybrid / On-site]\n\nThe full offer letter is attached. Please let us know your decision by [Deadline]. We are happy to answer any questions.\n\nWe look forward to having you on the team.\n\nBest regards,\n[Your name]\n[Company]"),
("Job offer","PT-PT","Final decision: hire","Proposta de trabalho – [Função] na [Empresa]",
"Olá [Nome],\n\nTemos o prazer de lhe apresentar uma proposta para a função de [Função] na [Empresa]!\n\nCondições principais:\n- Data de início: [Data]\n- Contrato: [Tipo]\n- Remuneração: [Valor] € brutos anuais\n- Modelo de trabalho: [Remoto / Híbrido / Presencial]\n\nSegue em anexo a carta de proposta completa. Pedimos que nos dê uma resposta até [Prazo]. Estamos disponíveis para qualquer dúvida.\n\nSerá um gosto tê-la/o na equipa.\n\nCom os melhores cumprimentos,\n[O seu nome]\n[Empresa]"),
]
root = Path("templates")
with (root/"email-templates.csv").open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["Template","Language","When to use","Subject","Body"]); w.writerows(T)
md = ["# Email templates", "", "Placeholders are in square brackets. EN and PT-PT versions of each message.", ""]
for name, lang, when, subj, body in T:
    md += [f"## {name} ({lang})", "", f"Use when: {when.lower()}.", "", f"**Subject:** {subj}", "", "```text", body, "```", ""]
(root/"email-templates.md").write_text("\n".join(md), encoding="utf-8")
