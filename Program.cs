using System;
using System.Drawing; //Serve para dar estilo ao texto. Ex: cor, nome, deixar negrito, etc...
using Spire.Doc;
using Spire.Doc.Documents;

namespace ExportWordPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Digite o titulo do arquivo");
            string title = Console.ReadLine();
            // Criando um novo documento com o nome de doc
            Document doc = new Document();

            // Criando uma seção dentro do doc
            // A cada seção criada uma nova página é adicionada
            Section secao = doc.AddSection();

            // Criando dois paragrado dentro da seção
            Paragraph titulo = secao.AddParagraph();
            Paragraph texto = secao.AddParagraph();

            // Insiro na minha variavel titulo algum valor
            // Ou seja => Hello World
            titulo.AppendText($"{title}\n");

            // Insiro na minha variavel texto algum valor
            texto.AppendText("Ai você fala o seguinte: '- Mas vocês acabaram isso?' Vou te falar: -'Não, está em andamento!' Tem obras que 'vai' durar pra depois de 2010. Agora, por isso, nós já não desenhamos, não começamos a fazer projeto do que nós 'podêmo fazê'? 11, 12, 13, 14... Por que é que não? Primeiro eu queria cumprimentar os internautas. -Oi Internautas! Depois dizer que o meio ambiente é sem dúvida nenhuma uma ameaça ao desenvolvimento sustentável. E isso significa que é uma ameaça pro futuro do nosso planeta e dos nossos países. O desemprego beira 20%, ou seja, 1 em cada 4 portugueses.A única área que eu acho, que vai exigir muita atenção nossa, e aí eu já aventei a hipótese de até criar um ministério. É na área de... Na área... Eu diria assim, como uma espécie de analogia com o que acontece na área agrícola.A população ela precisa da Zona Franca de Manaus, porque na Zona franca de Manaus, não é uma zona de exportação, é uma zona para o Brasil. Portanto ela tem um objetivo, ela evita o desmatamento, que é altamente lucravito. Derrubar arvores da natureza é muito lucrativo!\n\n");


            // ** Configurando o titulo do arquivo **
            
            // Alinhar o titulo na página centro da página
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            // Criando um novo estilo para o nosso paragrafo
            // (doc) é o nome do nosso documento
            ParagraphStyle estiloTitulo = new ParagraphStyle(doc);

            estiloTitulo.Name = "Cor do titulo"; // Define o nome da classe do estiloTitulo
            estiloTitulo.CharacterFormat.TextColor = Color.Red; //  Transforma a cor do texto em dark magenta
            estiloTitulo.CharacterFormat.Bold = true; // Transforma o texto em negrito

            estiloTitulo.CharacterFormat.FontSize = 25; // Transformando o tamnho da fonte

            estiloTitulo.CharacterFormat.FontName = "Arial"; // Mudadndo a fonte do documento
            // Adicionar o estilo e colocar como usavel no documento
            // doc => nome do documento
            // Add(Nome do estiloTitulo Criando)
            doc.Styles.Add(estiloTitulo);

            // Aqui aplica todos os estilos dados anteriormente 
            titulo.ApplyStyle(estiloTitulo.Name);


            // ** Fim da configuração do titulo **


            // ** Configurando o texto **

            texto.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            ParagraphStyle estiloTexto = new ParagraphStyle(doc);
            estiloTexto.Name = "Estilização dos texos";
            estiloTexto.CharacterFormat.FontSize = 12;
            estiloTexto.CharacterFormat.FontName = "Arial";

            // ** Fim da configuração do texto ** 

            // Exportando o arquivo para a pasta docs com o nome de exemplo

            doc.SaveToFile($@"docs\exemplo.docx", FileFormat.Docx);
            doc.SaveToFile($@"docs\exemplo.PDF", FileFormat.PDF);

        }
    }
}
