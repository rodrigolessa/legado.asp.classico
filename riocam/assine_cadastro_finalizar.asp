      <table width="100%" height="100%" border=0 cellspacing=0 cellpadding=0>
        <tr>
          <td colspan=2 valign="top">
            <table border=0 cellspacing=0 cellpadding=0 width=640>
<%
              if session("avulso_partic") = "P" then
%>
				<tr align="justify">
					<td align="justify">
						<b>
						INÍCIO DO PROCESSO DE INSCRIÇÃO.
						<p>Para a finalização do processo e efetivação da inscrição será necessário o pagamento da taxa de inscrição via Boleto Bancário no valor de R$ <%=valor%>.<br><br>
						* O boleto bancário ficará disponível até o último dia de inscrição.<br><br>
						** Em caso de feriado ou evento que acarrete o fechamento de agências bancárias, na localidade em que se encontra o profissional, o boleto deverá ser pago antecipadamente.<br><br>
						***As inscrições cujos pagamentos não forem efetuados até as 16 horas do último dia de inscrição serão desconsideradas.<br><br>
						****Não haverá restituição de valor seja qual for o motivo alegado.<br><br>
						<font color="red">Verifique se o seu computador possui bloqueador de pop-ups, pois este recurso impede a abertura da página do boleto até ser desativado.</font>
						<br><br>
						</p>
						</b>
					</td>
				</tr>
<%
              else
%>
                <tr align="justify">
                  <td align="justify">
                    <b>INÍCIO DO PROCESSO DE INSCRIÇÃO.</b>
                    <p>
                    Para a finalização do processo e efetivação da inscrição serão necessários:
                    <UL>
                    	<LI>O pagamento da taxa de inscrição via Boleto bancário no valor de R$ <%=valor%> e;</LI>
                    	<LI>A Declaração de Habilitação com firma reconhecida onde será atestado que o profissional não é vinculado a uma Instituição vinculada ao BACEN ou CVM.</LI>
<!--                    	<LI>Uma cópia autenticada da página da Carteira Profissional atrás da foto, que contém seus dados cadastrais; a página que contém a última admissão e o último desligamento (quando for o caso) e a página posterior em branco.</LI>-->
                    </UL>
                    * O boleto bancário ficará disponível até o último dia de inscrição.<br><br>
                    ** A inscrição será desconsiderada se: o pagamento do boleto bancário não for efetuado e/ou a declaração de habilitação não for encaminhada a ANBID até o último dia de inscrição.<br><br>
                    *** Em caso de feriado ou evento que acarrete o fechamento de agências bancárias, na localidade em que se encontra o profissional, o boleto deverá ser pago antecipadamente. Não haverá restituição do valor pago, seja qual for o motivo alegado.<br><br>
					<font color="red">Verifique se o seu computador possui bloqueador de pop-ups, pois este recurso impede a abertura da página do boleto até ser desativado.</font>
                    </p>
                  </td>
                </tr>
<%
              end if
%>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><input type="button" name="btnProcessar" value="Emitir Boleto" class="botao" onclick="document.all['pBoleto'].Document.location.href='/processaBoletoAvulso.asp?codprova=<%=session("avulso_codprova")%>'"></td>
                <td>&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
<!--Fim do tabelão-->
<!--iframe PARA IMPRESSÃO DO BOLETO BANCÁRIO src="teste\processaBoleto.asp...visibity:hidden;" -->
<iframe name="pBoleto" id="pBoleto" style="visibility:hidden;width:0px;height:0px"></iframe>
<!--<iframe name="pBoleto" id="pBoleto" style="width:500px;height:300px"></iframe>-->
