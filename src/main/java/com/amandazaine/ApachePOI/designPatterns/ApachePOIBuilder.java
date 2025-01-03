package com.amandazaine.ApachePOI.designPatterns;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//Implementa o design pattern: builder
public class ApachePOIBuilder {

    public static class CellStyleBuilder {
        private short corPadrao;
        private XSSFColor corPersonalizada;
        private FillPatternType tipoPreenchimento;
        private XSSFFont fonte;
        private String formatoDadoCelula;
        private HorizontalAlignment alinhamentoHorizontal;
        private VerticalAlignment alinhamentoVertical;
        private BorderStyle estiloBordaSuperior;
        private BorderStyle estiloBordaInferior;
        private BorderStyle estiloBordaEsquerda;
        private BorderStyle estiloBordaDireita;

        public CellStyleBuilder setCorPadrao(short corPadrao) {
            this.corPadrao = corPadrao;
            return this;
        }

        public CellStyleBuilder setCorPersonalizada(String corPersonalizada) {
            try {
                byte[] rgb = Hex.decodeHex(corPersonalizada);
                this.corPersonalizada = new XSSFColor(rgb);
                return this;
            } catch (DecoderException e) {
                throw new RuntimeException("Erro ao decodificar a cor: " + e);
            }
        }

        public CellStyleBuilder setTipoPreenchimento(FillPatternType tipoPreenchimento) {
            this.tipoPreenchimento = tipoPreenchimento;
            return this;
        }

        public CellStyleBuilder setFonte(XSSFFont fonte) {
            this.fonte = fonte;
            return this;
        }

        public CellStyleBuilder setFormatoDadoCelula(String formatoDadoCelula) {
            this.formatoDadoCelula = formatoDadoCelula;
            return this;
        }

        public CellStyleBuilder setAlinhamentoHorizontal(HorizontalAlignment alinhamentoHorizontal) {
            this.alinhamentoHorizontal = alinhamentoHorizontal;
            return this;
        }

        public CellStyleBuilder setAlinhamentoVertical(VerticalAlignment alinhamentoVertical) {
            this.alinhamentoVertical = alinhamentoVertical;
            return this;
        }

        public CellStyleBuilder setEstiloBordaSuperior(BorderStyle estiloBordaSuperior) {
            this.estiloBordaSuperior = estiloBordaSuperior;
            return this;
        }

        public CellStyleBuilder setEstiloBordaInferior(BorderStyle estiloBordaInferior) {
            this.estiloBordaInferior = estiloBordaInferior;
            return this;
        }

        public CellStyleBuilder setEstiloBordaEsquerda(BorderStyle estiloBordaEsquerda) {
            this.estiloBordaEsquerda = estiloBordaEsquerda;
            return this;
        }

        public CellStyleBuilder setEstiloBordaDireita(BorderStyle estiloBordaDireita) {
            this.estiloBordaDireita = estiloBordaDireita;
            return this;
        }

        public XSSFCellStyle build(XSSFWorkbook workbook) {
            XSSFCellStyle estiloCelula = workbook.createCellStyle();

            if(this.corPadrao != 0 ) {
                estiloCelula.setFillForegroundColor(corPadrao);
            }

            if(this.corPersonalizada != null) {
                estiloCelula.setFillForegroundColor(corPersonalizada);
            }

            if(this.tipoPreenchimento != null) {
                estiloCelula.setFillPattern(tipoPreenchimento);
            }

            if(this.fonte != null) {
                estiloCelula.setFont(fonte);
            }

            if(this.formatoDadoCelula != null && !this.formatoDadoCelula.isBlank()) {
                estiloCelula.setDataFormat(workbook.createDataFormat().getFormat(formatoDadoCelula));
            }

            if(this.alinhamentoHorizontal != null) {
                estiloCelula.setAlignment(alinhamentoHorizontal);
            }

            if(this.alinhamentoVertical != null) {
                estiloCelula.setVerticalAlignment(alinhamentoVertical);
            }

            if(this.estiloBordaSuperior != null) {
                estiloCelula.setBorderTop(estiloBordaSuperior);
            }

            if(this.estiloBordaInferior != null) {
                estiloCelula.setBorderBottom(estiloBordaInferior);
            }

            if(this.estiloBordaEsquerda != null) {
                estiloCelula.setBorderLeft(estiloBordaEsquerda);
            }

            if(this.estiloBordaDireita != null) {
                estiloCelula.setBorderRight(estiloBordaDireita);
            }

            return estiloCelula;
        }
    }

    public static class FontBuilder {
        private String nomeFont;
        private boolean negrito = false;
        private boolean italico = false;
        private short tamanhoFont;
        private short corPadrao;
        private XSSFColor corPersonalizada;
        private FontUnderline underline;

        public FontBuilder setNomeFont(String nomeFont) {
            this.nomeFont = nomeFont;
            return this;
        }

        public FontBuilder setNegrito(boolean negrito) {
            this.negrito = negrito;
            return this;
        }

        public FontBuilder setItalico(boolean italico) {
            this.italico = italico;
            return this;
        }

        public FontBuilder setTamanhoFont(short tamanhoFont) {
            this.tamanhoFont = tamanhoFont;
            return this;
        }

        public FontBuilder setCorPadrao(short corPadrao) {
            this.corPadrao = corPadrao;
            return this;
        }

        public FontBuilder setcorPersonalizada(String corPersonalizada) {
            try {
                byte[] rgb = Hex.decodeHex(corPersonalizada);
                this.corPersonalizada = new XSSFColor(rgb);
                return this;
            } catch (DecoderException e) {
                throw new RuntimeException("Erro ao decodificar a cor: " + e);
            }
        }

        public FontBuilder setUnderline(FontUnderline underline) {
            this.underline = underline;
            return this;
        }

        public XSSFFont build(XSSFWorkbook workbook) {
            XSSFFont fonte = workbook.createFont();

            if(nomeFont != null) {
                fonte.setFontName(nomeFont);
            }

            if(negrito) {
                fonte.setBold(negrito);
            }

            if(italico) {
                fonte.setItalic(italico);
            }

            if(tamanhoFont != 0) {
                fonte.setFontHeight(tamanhoFont);
            }

            if(corPadrao != 0) {
                fonte.setColor(corPadrao);
            }

            if(corPersonalizada != null) {
                fonte.setColor(corPersonalizada);
            }

            if(underline != null) {
                fonte.setUnderline(underline);
            }

            return fonte;
        }
    }
}
