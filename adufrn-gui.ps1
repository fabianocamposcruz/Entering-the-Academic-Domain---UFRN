# Script PowerShell para ingressar computadores desktop com Windows no domínio acadêmico da UFRN (ad.ufrn.br)
# Instruções sobre execução consultar o arquivo <Readme.txt>
# ATENÇÃO!!! SEMPRE VERIFICAR SE O SETOR/UNIDADE ONDE SERÁ APLICADO ESTÁ ATUALIZADO!!
# Desenvolvido por Fabiano Campos
# Versão 12/11/2025


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ================= DADOS ESTÁTICOS =================
$codigosUnidade = @(
    [pscustomobject]@{Codigo="1415"; Nome="DEPARTAMENTO DE ARQUITETURA"}
    [pscustomobject]@{Codigo="1416"; Nome="DEPARTAMENTO DE ENGENHARIA DE MATERIAIS"}
    [pscustomobject]@{Codigo="1417"; Nome="DEPARTAMENTO DE ENGENHARIA CIVIL E AMBIENTAL"}
    [pscustomobject]@{Codigo="1418"; Nome="DEPARTAMENTO DE ENGENHARIA DE COMPUTACAO E AUTOMACAO"}
    [pscustomobject]@{Codigo="1419"; Nome="DEPARTAMENTO DE ENGENHARIA ELETRICA"}
    [pscustomobject]@{Codigo="1420"; Nome="DEPARTAMENTO DE ENGENHARIA MECANICA"}
    [pscustomobject]@{Codigo="1421"; Nome="DEPARTAMENTO DE ENGENHARIA QUIMICA"}
    [pscustomobject]@{Codigo="1422"; Nome="DEPARTAMENTO DE ENGENHARIA PRODUCAO"}
    [pscustomobject]@{Codigo="1424"; Nome="DEPARTAMENTO DE ENGENHARIA TEXTIL"}
    [pscustomobject]@{Codigo="1431"; Nome="DIRECAO DO CENTRO DE TECNOLOGIA"}
    [pscustomobject]@{Codigo="1433"; Nome="DEPARTAMENTO DE ENGENHARIA DE PETROLEO"}
    [pscustomobject]@{Codigo="1435"; Nome="DEPARTAMENTO DE ENGENHARIA BIOMEDICA"}
    [pscustomobject]@{Codigo="1436"; Nome="DEPARTAMENTO DE ENGENHARIA DE COMUNICACOES"}
)
$locais = @(
    [pscustomobject]@{Codigo = "CT"; Nome = "CT"}
    [pscustomobject]@{Codigo = "CTEC"; Nome = "CTEC"}
    [pscustomobject]@{Codigo = "Galinheiro"; Nome = "GALINHEIRO"}
    [pscustomobject]@{Codigo = "LARHISA"; Nome = "LARHISA"}
    [pscustomobject]@{Codigo = "NEAU"; Nome = "NEAU"}
    [pscustomobject]@{Codigo = "NTI"; Nome = "NTI"}
    [pscustomobject]@{Codigo = "PGTEC"; Nome = "PGTEC"}
    [pscustomobject]@{Codigo = "SetorIV"; Nome = "SETOR IV"}
)
$setoresPorLocal = @{
    "CT" = @(
        @{Codigo = "PIT"; Nome = "Pitagoras"},
        @{Codigo = "DIR"; Nome = "Diretoria"}       
    )
    "CTEC" = @(
        @{Codigo = "LPO"; Nome = "Laboratorio de Pesquisa Operacional - DEP"},
        @{Codigo = "LMA"; Nome = "Laboratorio de Modelagem Ambiental - LMA"}        
    )
    "Galinheiro" = @(
        @{Codigo = 1250; Nome = "AINDA SEM CADASTRO"},
        @{Codigo = 1251; Nome = "AINDA SEM CADASTRO"}
    )
    "LARHISA" = @(
        @{Codigo = 1260; Nome = "AINDA SEM CADASTRO"},
        @{Codigo = 1261; Nome = "AINDA SEM CADASTRO"}       
    )
    "NEAU" = @(
        @{Codigo = 1270; Nome = "AINDA SEM CADASTRO"},
        @{Codigo = 1271; Nome = "AINDA SEM CADASTRO"}
    )
    "NTI" = @(
        @{Codigo = "NP1"; Nome = "NUPEG I"},
        @{Codigo = "NP2"; Nome = "NUPEG II"},
	@{Codigo = "LMS"; Nome = "Laboratorio de Modelagem e Simulacao - DEQ"},
        @{Codigo = "LET"; Nome = "Laboratorio de Engenharia Textil II - DET"}
    )
    "PGTEC" = @(
        @{Codigo = 1290; Nome = "AINDA SEM CADASTRO"},
        @{Codigo = 1291; Nome = "AINDA SEM CADASTRO"}
    )
    "SetorIV" = @(
        @{Codigo = "E04"; Nome = "Sl_Estudos"},
        @{Codigo = "I01"; Nome = "Administracao"},
        @{Codigo = "B03"; Nome = "Auditorio"},
		@{Codigo = "OPT"; Nome = "Ensino"},
		@{Codigo = "H01"; Nome = "H1"},
		@{Codigo = "C06"; Nome = "LAB 1 - C6"},
		@{Codigo = "D04"; Nome = "LAB 2 - D4"},
		@{Codigo = "E03"; Nome = "LAB 3 - E3"},
		@{Codigo = "E06"; Nome = "LAB 4 - E6"},
		@{Codigo = "G06"; Nome = "LAB 6 - G6"},
		@{Codigo = "E04"; Nome = "Supervisao"},
		@{Codigo = "E02"; Nome = "Suporte"}
    )
}

$serial = Get-WmiObject win32_bios | Select-Object -ExpandProperty SerialNumber
$dominio = Get-ComputerInfo | Select-Object -ExpandProperty CsDomain
$global:aquisicao = 0

if ($dominio -eq "ad.ufrn.br") {
	[System.Windows.Forms.MessageBox]::Show("O COMPUTADOR JÁ ESTÁ NO DOMÍNIO '$dominio'. NENHUMA ALTERAÇÃO É NECESSÁRIA.", "INFORMAÇÃO", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    exit
}
else {
	if (Test-Path -Path "C:\$Env:HOMEPATH\validaufrn.txt") {
		$setor = Get-content -path "C:\$Env:HOMEPATH\variavel_setor.txt"
		$unidade = Get-content -path "C:\$Env:HOMEPATH\variavel_unidade.txt"
		Remove-Item -Path "C:\$Env:HOMEPATH\variavel_setor.txt" -Force
		Remove-Item -Path "C:\$Env:HOMEPATH\variavel_unidade.txt" -Force
		Remove-Item -Path "C:\$Env:HOMEPATH\validaufrn.txt" -Force
		Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "adufrn" -ErrorAction SilentlyContinue
		$hostname = (Get-WmiObject Win32_ComputerSystem).Name
		$resposta3 = [System.Windows.Forms.MessageBox]::Show("DESEJA INGRESSAR NO DOMÍNIO: 'ad.ufrn.br'? SIM, PARA INSERIR CREDENCIAIS. NÃO, PARA SAIR.", "CONFIRMAÇÃO", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
		if ($resposta3 -eq [System.Windows.Forms.DialogResult]::Yes){
			Add-Computer -ComputerName $hostname -DomainName "ad.ufrn.br" -Credential (Get-Credential) -OUPath "OU=$setor,OU=$unidade,OU=COMPUTADORES,OU=CT,OU=UFRN,DC=ad,DC=ufrn,DC=br" -Restart
		} else {
			exit
		}
	}
	else {
		# Formata a janela
		$main_form = New-Object System.Windows.Forms.Form
		$main_form.Text ='UFRN/CT: INGRESSAR WINDOWS NO DOMÍNIO ACADÊMICO'
		$main_form.Width = 580
		$main_form.Height = 500
		$main_form.FormBorderStyle = 'FixedDialog'
		$main_form.MaximizeBox = $false
		$main_form.MinimizeBox = $false
		$main_form.StartPosition = "CenterScreen"
		$main_form.BackColor = 'CornflowerBlue'
		
		$LabelDescricao = New-Object System.Windows.Forms.Label
		$LabelDescricao.Text = "Software desenvolvido em PowerShell para ingresso de computadores Windows `nsob gerência do CT no domínio acadêmico da UFRN: 'ad.ufrn.br'. `nCertifique-se de: 1) Estar logado com perfil de administrador; `n                              2) Este executável esteja em 'C:\Windows\System32\'."
		$LabelDescricao.Location  = New-Object System.Drawing.Point(80,20)
		$LabelDescricao.AutoSize = $true
		$LabelDescricao.ForeColor = 'White'
		$main_form.Controls.Add($LabelDescricao)
		
		$LabelRodape = New-Object System.Windows.Forms.Label
		$LabelRodape.Text = "©2025 - Desenvolvido por Pitágoras/CT"
		$LabelRodape.Location  = New-Object System.Drawing.Point(180,445)
		$LabelRodape.AutoSize = $true
		$LabelRodape.ForeColor = 'White'
		$main_form.Controls.Add($LabelRodape)
		
		$LabelUnidade = New-Object System.Windows.Forms.Label
		$LabelUnidade.Text = "Unidade vinculada: "
		$LabelUnidade.Location  = New-Object System.Drawing.Point(10,95)
		$LabelUnidade.ForeColor = 'White'
		$LabelUnidade.AutoSize = $true
		$main_form.Controls.Add($LabelUnidade)

		$ComboBoxUnidade = New-Object System.Windows.Forms.ComboBox
		$ComboBoxUnidade.Width = 350
		$ComboBoxUnidade.Location  = New-Object System.Drawing.Point(120,90)
		$ComboBoxUnidade.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
		$ComboBoxUnidade.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
		$ComboBoxUnidade.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
		$codigosUnidade | ForEach-Object { $ComboBoxUnidade.Items.Add($_.Nome) | Out-Null}
		$main_form.Controls.Add($ComboBoxUnidade)

		# Label para exibir o código
		$LabelCodigo = New-Object System.Windows.Forms.Label
		$LabelCodigo.Text = "Código: "
		$LabelCodigo.Location = New-Object System.Drawing.Point(10, 125)
		$LabelCodigo.ForeColor = 'White'
		$LabelCodigo.AutoSize = $true
		$main_form.Controls.Add($LabelCodigo)
		$ComboBoxUnidade.Add_SelectedIndexChanged({
			$index = $ComboBoxUnidade.SelectedIndex
			if ($index -ge 0) {
				$global:codigo_unidade = $codigosUnidade[$index].Codigo
				$LabelCodigo.Text = "Código: $global:codigo_unidade"
			}
		})

		# ----- Localização -----
		$LabelLocal = New-Object System.Windows.Forms.Label
		$LabelLocal.Text = "Localização:"
		$LabelLocal.Location  = New-Object System.Drawing.Point(10, 155)
		$LabelLocal.ForeColor = 'White'
		$LabelLocal.AutoSize = $true
		$main_form.Controls.Add($LabelLocal)

		$ComboBoxLocal = New-Object System.Windows.Forms.ComboBox
		$ComboBoxLocal.Width = 350
		$ComboBoxLocal.Location = New-Object System.Drawing.Point(120, 150)
		$ComboBoxLocal.AutoCompleteMode = 'SuggestAppend'
		$ComboBoxLocal.AutoCompleteSource = 'ListItems'
		$locais | ForEach-Object { $ComboBoxLocal.Items.Add($_.Nome) |Out-Null }
		$main_form.Controls.Add($ComboBoxLocal)

		# ----- Setor/Lab/Sala -----
		$LabelSetor = New-Object System.Windows.Forms.Label
		$LabelSetor.Text = "Setor/Lab/Sala:"
		$LabelSetor.Location  = New-Object System.Drawing.Point(10, 185)
		$LabelSetor.ForeColor = 'White'
		$LabelSetor.AutoSize = $true
		$main_form.Controls.Add($LabelSetor)

		$ComboBoxSetor = New-Object System.Windows.Forms.ComboBox
		$ComboBoxSetor.Width = 350
		$ComboBoxSetor.Location = New-Object System.Drawing.Point(120, 180)
		$ComboBoxSetor.AutoCompleteMode = 'SuggestAppend'
		
		$main_form.Controls.Add($ComboBoxSetor)

		# Atualiza ComboBox de setores com base no local
		$ComboBoxLocal.Add_SelectedIndexChanged({
			$ComboBoxSetor.Items.Clear()
			$ComboBoxSetor.Text = ""

			$localIndex = $ComboBoxLocal.SelectedIndex
			if ($localIndex -ge 0) {
				$localSelecionado = $locais[$localIndex]
				$global:localCodigoSelecionado = $localSelecionado.Codigo

				$setores = $setoresPorLocal[$global:localCodigoSelecionado]
				foreach ($setor in $setores) {
					$ComboBoxSetor.Items.Add($setor.Nome) | Out-Null
				}
			}
		})

		# Quando selecionar setor, registra nome e código
		$ComboBoxSetor.Add_SelectedIndexChanged({
			$global:setorNomeSelecionado = $ComboBoxSetor.SelectedItem

			if ($localCodigoSelecionado -and $setoresPorLocal.ContainsKey($localCodigoSelecionado)) {
				$setor = $setoresPorLocal[$localCodigoSelecionado] | Where-Object { $_.Nome -eq $setorNomeSelecionado }

				if ($setor) {
					$global:setorCodigoSelecionado = $setor.Codigo
				}
			} 
		})

		# -------------------- Grupo: Identificação Única --------------------
		$GroupBoxIdent = New-Object System.Windows.Forms.GroupBox
		$GroupBoxIdent.Text = "Identificação única:"
		$GroupBoxIdent.ForeColor = 'White'
		$GroupBoxIdent.Location = New-Object System.Drawing.Point(10, 210)
		$GroupBoxIdent.Size = New-Object System.Drawing.Size(550, 80)
		$main_form.Controls.Add($GroupBoxIdent)

		$RadioPatrimonio = New-Object System.Windows.Forms.RadioButton
		$RadioPatrimonio.Text = "Patrimônio (SIPAC)"
		$RadioPatrimonio.Location = New-Object System.Drawing.Point(10, 20)
		$RadioPatrimonio.Checked =$true
		$RadioPatrimonio.AutoSize = $true
		$GroupBoxIdent.Controls.Add($RadioPatrimonio)

		$RadioSerial = New-Object System.Windows.Forms.RadioButton
		$RadioSerial.Text = "Serial Number"
		$RadioSerial.Location = New-Object System.Drawing.Point(200, 20)
		$RadioSerial.AutoSize = $true
		$GroupBoxIdent.Controls.Add($RadioSerial)

		$TextBoxPatrimonio = New-Object System.Windows.Forms.TextBox
		$TextBoxPatrimonio.Width = 80
		$TextBoxPatrimonio.Location = New-Object System.Drawing.Point(10, 45)
		$TextBoxPatrimonio.Enabled = $true
		$GroupBoxIdent.Controls.Add($TextBoxPatrimonio)
		# Evento para restringir a entrada apenas a números e 4 caracteres
		$TextBoxPatrimonio.Add_KeyPress({
			param($sender, $e)
			# Impede entrada se não for número ou se já tiver 4 caracteres
			if ($e.KeyChar -match '\D' -and $e.KeyChar -ne [char]8) {
				$e.Handled = $true
			} elseif ($TextBoxPatrimonio.Text.Length -ge 4 -and $e.KeyChar -ne [char]8) {
				$e.Handled = $true
			}
		})
		# Atualiza a variável global quando o texto for alterado
		$TextBoxPatrimonio.Add_TextChanged({
			if ($RadioPatrimonio.Checked -and $TextBoxPatrimonio.Text.Length -eq 4) {
				$global:identificacao = $TextBoxPatrimonio.Text
			} 
		})
		# Eventos para identificação única
		$RadioPatrimonio.Add_CheckedChanged({
			if ($RadioPatrimonio.Checked) {
				$TextBoxPatrimonio.Enabled = $true
				$TextBoxPatrimonio.Focus()
			}
		})

		$RadioSerial.Add_CheckedChanged({
			if ($RadioSerial.Checked) {
				$TextBoxPatrimonio.Enabled = $false
				$TextBoxPatrimonio.Text = ""
				$global:identificacao = $serial.Substring($serial.Length - 4)
				
			}
		})

		$TextBoxPatrimonio.Add_TextChanged({
			if ($RadioPatrimonio.Checked -and $TextBoxPatrimonio.Text.Length -eq 4) {
				$global:identificacao = $TextBoxPatrimonio.Text
			}
		})

		# -------------------- Grupo: Origem da Aquisição --------------------
		$GroupBoxAquisicao = New-Object System.Windows.Forms.GroupBox
		$GroupBoxAquisicao.Text = "Origem da aquisição:"
		$GroupBoxAquisicao.ForeColor = 'White'
		$GroupBoxAquisicao.Location = New-Object System.Drawing.Point(10, 300)
		$GroupBoxAquisicao.Size = New-Object System.Drawing.Size(550, 60)
		$main_form.Controls.Add($GroupBoxAquisicao)

		$RadioAquisicaoUFRN = New-Object System.Windows.Forms.RadioButton
		$RadioAquisicaoUFRN.Text = "Aquisição da UFRN"
		$RadioAquisicaoUFRN.Location = New-Object System.Drawing.Point(10, 20)
		$RadioAquisicaoUFRN.Checked =$true
		$RadioAquisicaoUFRN.AutoSize = $true
		$GroupBoxAquisicao.Controls.Add($RadioAquisicaoUFRN)

		$RadioProjetoPesquisa = New-Object System.Windows.Forms.RadioButton
		$RadioProjetoPesquisa.Text = "Projeto de Pesquisa"
		$RadioProjetoPesquisa.Location = New-Object System.Drawing.Point(200, 20)
		$RadioProjetoPesquisa.AutoSize = $true
		$GroupBoxAquisicao.Controls.Add($RadioProjetoPesquisa)

		# Eventos para aquisição
		$RadioAquisicaoUFRN.Add_CheckedChanged({
			if ($RadioAquisicaoUFRN.Checked) {
				$global:aquisicao = 0
			}
		})

		$RadioProjetoPesquisa.Add_CheckedChanged({
			if ($RadioProjetoPesquisa.Checked) {
				$global:aquisicao = 1
			}
		})

		# ---------------- Botão: Renomear computador e reiniciar ----------------
		$ButtonRenomear = New-Object System.Windows.Forms.Button
		$ButtonRenomear.Text = "RENOMEAR COMPUTADOR E INSERIR NO DOMÍNIO"
		$ButtonRenomear.ForeColor = 'White'
		$ButtonRenomear.Location = New-Object System.Drawing.Point(85, 375)
		$ButtonRenomear.Size = New-Object System.Drawing.Size(400, 50)

		$ButtonRenomear.Add_Click({
			# Verificar se todos os campos obrigatórios estão preenchidos
			if (-not $codigo_unidade -or -not $localCodigoSelecionado -or -not $identificacao -or $aquisicao -or $global:setorCodigoSelecionado -eq $null) {
				[System.Windows.Forms.MessageBox]::Show("POR FAVOR, PREENCHA TODOS OS CAMPOS OBRIGATÓRIOS.", "ATENÇÃO", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				return				
			}

			# Define 0 ou 1 de acordo com o tipo de identificação, Patrimônio ou Serial, respectivamente
			$select_identificacao = if ($RadioPatrimonio.Checked) { "0" } else { "1" }

			# Cria o novo nome do computador
			$new_name = "$codigo_unidade$setorCodigoSelecionado$select_identificacao$identificacao" + "D" + "$aquisicao" + "W"

			try {
				# Altera os endereços DNS
				Get-WmiObject -Class Win32_IP4RouteTable | Where-Object { $_.destination -eq '0.0.0.0' -and $_.mask -eq '0.0.0.0'} |
					Sort-Object metric1 | Select-Object -First 1 -ExpandProperty interfaceindex |
					ForEach-Object {
						Set-DnsClientServerAddress -InterfaceIndex $_ -ServerAddresses ('10.3.158.13','10.3.158.14')
					}

				# Verificar se o nome do computador já é o mesmo
				$current_hostname = (Get-WmiObject Win32_ComputerSystem).Name
				if ($current_hostname -eq $new_name) {
					$resposta1 = [System.Windows.Forms.MessageBox]::Show("O NOME DO COMPUTADOR JÁ É: '$new_name'. NENHUMA ALTERAÇÃO É NECESSÁRIA. DESEJA INGRESSAR NO DOMÍNIO? SIM, PARA INSERIR CREDENCIAIS. NÃO, PARA SAIR.", "CONFIRMAÇÃO", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
					if ($resposta1 -eq [System.Windows.Forms.DialogResult]::Yes){
						Add-Computer -ComputerName $current_hostname -DomainName "ad.ufrn.br" -Credential (Get-Credential) -OUPath "OU= $setorNomeSelecionado,OU= $localCodigoSelecionado,OU=COMPUTADORES,OU=CT,OU=UFRN,DC=ad,DC=ufrn,DC=br" -Restart
					} else {
						return
					} 
					
				}
				else {
					# Confirma nome gerado e pergunta se deseja continuar
					$resposta2 = [System.Windows.Forms.MessageBox]::Show("O COMPUTADOR SERA RENOMEADO PARA: `n$new_name`nDESEJA CONFIRMAR E REINICIAR?", "CONFIRMAÇÃO", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
					if ($resposta2 -eq [System.Windows.Forms.DialogResult]::Yes) {
						$validacaoPath = "C:\$env:HOMEPATH\validaufrn.txt"
						New-Item -Path $validacaoPath -ItemType File -Force
						$setorNomeSelecionado | Out-File -FilePath "C:\$env:HOMEPATH\variavel_setor.txt" -Encoding utf8
						$localCodigoSelecionado | Out-File -FilePath "C:\$env:HOMEPATH\variavel_unidade.txt" -Encoding utf8
						$scriptPath = "C:\Windows\System32\adufrn.exe"
						#$task = "powershell.exe -ExecutionPolicy Bypass -File `"$scriptPath`""
						$task = $scriptPath
						Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "adufrn" -Value $task
						Rename-Computer -NewName $new_name -Restart
					} else { 
							return
						}
				}
			} catch {
				[System.Windows.Forms.MessageBox]::Show("Ocorreu um erro: $($_.Exception.Message)", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		})


		$main_form.Controls.Add($ButtonRenomear)
		$main_form.ShowDialog() | Out-Null		
	}
}

