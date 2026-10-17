# epub2pdf by Jair Lima

Converte um e-book EPUB em PDF no formato A5, com margens espelhadas, página de rosto e sumário paginado, para impressão de livro físico.

## Como funciona

O script monta um DOCX temporário a partir do EPUB e usa a automação COM do Microsoft Word para paginar, atualizar o sumário e exportar o PDF. Por isso:

- **Requer Windows e o Microsoft Word instalado.**
- O arquivo intermediário é apagado ao final.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```bash
python epub2pdf.py livro.epub livro.pdf
python epub2pdf.py livro.epub livro.pdf --font "Garamond" --size 11
```

Opções: `--font` (padrão Georgia), `--size` (10,5), `--margin-inner` (2,2 cm), `--margin-outer` (1,8 cm), `--margin-top` e `--margin-bottom` (1,8 cm).

## Especificação

Detalhes em [SPEC.md](SPEC.md).

## Licença

MIT, veja [LICENSE](LICENSE).
