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
						IN�CIO DO PROCESSO DE INSCRI��O.
						<p>Para a finaliza��o do processo e efetiva��o da inscri��o ser� necess�rio o pagamento da taxa de inscri��o via Boleto Banc�rio no valor de R$ <%=valor%>.<br><br>
						* O boleto banc�rio ficar� dispon�vel at� o �ltimo dia de inscri��o.<br><br>
						** Em caso de feriado ou evento que acarrete o fechamento de ag�ncias banc�rias, na localidade em que se encontra o profissional, o boleto dever� ser pago antecipadamente.<br><br>
						***As inscri��es cujos pagamentos n�o forem efetuados at� as 16 horas do �ltimo dia de inscri��o ser�o desconsideradas.<br><br>
						****N�o haver� restitui��o de valor seja qual for o motivo alegado.<br><br>
						<font color="red">Verifique se o seu computador possui bloqueador de pop-ups, pois este recurso impede a abertura da p�gina do boleto at� ser desativado.</font>
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
                    <b>IN�CIO DO PROCESSO DE INSCRI��O.</b>
                    <p>
                    Para a finaliza��o do processo e efetiva��o da inscri��o ser�o necess�rios:
                    <UL>
                    	<LI>O pagamento da taxa de inscri��o via Boleto banc�rio no valor de R$ <%=valor%> e;</LI>
                    	<LI>A Declara��o de Habilita��o com firma reconhecida onde ser� atestado que o profissional n�o � vinculado a uma Institui��o vinculada ao BACEN ou CVM.</LI>
<!--                    	<LI>Uma c�pia autenticada da p�gina da Carteira Profissional atr�s da foto, que cont�m seus dados cadastrais; a p�gina que cont�m a �ltima admiss�o e o �ltimo desligamento (quando for o caso) e a p�gina posterior em branco.</LI>-->
                    </UL>
                    * O boleto banc�rio ficar� dispon�vel at� o �ltimo dia de inscri��o.<br><br>
                    ** A inscri��o ser� desconsiderada se: o pagamento do boleto banc�rio n�o for efetuado e/ou a declara��o de habilita��o n�o for encaminhada a ANBID at� o �ltimo dia de inscri��o.<br><br>
                    *** Em caso de feriado ou evento que acarrete o fechamento de ag�ncias banc�rias, na localidade em que se encontra o profissional, o boleto dever� ser pago antecipadamente. N�o haver� restitui��o do valor pago, seja qual for o motivo alegado.<br><br>
					<font color="red">Verifique se o seu computador possui bloqueador de pop-ups, pois este recurso impede a abertura da p�gina do boleto at� ser desativado.</font>
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
<!--Fim do tabel�o-->
<!--iframe PARA IMPRESS�O DO BOLETO BANC�RIO src="teste\processaBoleto.asp...visibity:hidden;" -->
<iframe name="pBoleto" id="pBoleto" style="visibility:hidden;width:0px;height:0px"></iframe>
<!--<iframe name="pBoleto" id="pBoleto" style="width:500px;height:300px"></iframe>-->
