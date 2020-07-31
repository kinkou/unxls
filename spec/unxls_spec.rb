# frozen_string_literal: true

RSpec.describe Unxls do
  it 'Has a version number' do
    expect(Unxls::VERSION).to_not eq nil
  end

  it 'Only opens BIFF8 files for now' do
    expect {
      Unxls.parse(testfile('other', 'biff2-empty.xls'))
    }.to raise_error 'Sorry, BIFF2 is not supported yet'

    expect {
      Unxls.parse(testfile('other', 'biff3-empty.xls'))
    }.to raise_error 'Sorry, BIFF3 is not supported yet'

    expect {
      Unxls.parse(testfile('other', 'biff4-empty.xls'))
    }.to raise_error 'Sorry, BIFF4 is not supported yet'

    expect {
      Unxls.parse(testfile('other', 'biff5-empty.xls'))
    }.to raise_error 'Sorry, BIFF5 is not supported yet'
  end

  context 'Workbook stream, Globals substream' do
    context 'Single records' do
      before(:context) do
        @browser1 = Unxls::Biff8::Browser.new(testfile('biff8', 'empty.xls'))
        @browser2 = Unxls::Biff8::Browser.new(testfile('biff8', 'preferences.xls'))
      end

      it 'Decodes BOF record' do
        expect(@browser1.globals[:BOF][:dt_d]).to eq :globals
      end

      it 'Decodes CalcPrecision record' do
        expect(@browser1.globals[:CalcPrecision][:fFullPrec]).to eq false
        expect(@browser2.globals[:CalcPrecision][:fFullPrec]).to eq true
      end

      it 'Decodes CodePage record' do
        expect(@browser1.globals[:CodePage][:cv_d]).to eq :'UTF-16LE'
      end

      it 'Decodes Country record' do
        c1 = @browser1.globals[:Country]
        expect(c1[:iCountryDef]).to eq 1
        expect(c1[:iCountryDef_d]).to eq :"United States"
        expect(c1[:iCountryWinIni]).to eq 1
        expect(c1[:iCountryWinIni_d]).to eq :"United States"

        c2 = @browser2.globals[:Country]
        expect(c2[:iCountryDef]).to eq 81
        expect(c2[:iCountryDef_d]).to eq :Japan
        expect(c2[:iCountryWinIni]).to eq 7
        expect(c2[:iCountryWinIni_d]).to eq :Russia
      end

      it 'Decodes Date1904 record' do
        expect(@browser1.globals[:Date1904][:f1904DateSystem]).to eq false
        expect(@browser2.globals[:Date1904][:f1904DateSystem]).to eq true
      end

      it 'Decodes Theme record' do
        expect(@browser1.globals[:Theme]).to eq nil

        browser = Unxls::Biff8::Browser.new(testfile('biff8', 'font.xls'))
        theme = browser.globals[:Theme]

        expect(theme[:rgb_d].size).to eq 5
        filename = 'theme/theme/theme1.xml'
        expect(theme[:rgb_d].keys).to include filename

        require 'rexml/document'
        parsed_theme = REXML::Document.new(theme[:rgb_d][filename])
        expect(parsed_theme.root.elements['a:themeElements/a:clrScheme[@name="Office"]'].size).to eq 12
        expect(parsed_theme.root.elements['a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr'].value).to eq '000000'
        expect(parsed_theme.root.elements['a:themeElements/a:fontScheme[@name="Office"]'].size).to eq 2
        expect(parsed_theme.root.elements['a:themeElements/a:fmtScheme[@name="Office"]'].size).to eq 4
      end
    end

    context 'BoundSheet8 record' do
      before(:context) do
        @browser1 = Unxls::Biff8::Browser.new(testfile('biff8', 'empty.xls'))
        @browser2 = Unxls::Biff8::Browser.new(testfile('biff8', 'preferences.xls'))
      end

      specify 'For worksheet "Worksheet"' do
        bs = @browser2.globals[:BoundSheet8][0]
        expect(bs[:lbPlyPos]).to eq(@browser2.wbs[1][:BOF][:_record][:pos])
        expect(bs[:hsState]).to eq 0
        expect(bs[:hsState_d]).to eq :visible
        expect(bs[:dt]).to eq 0
        expect(bs[:dt_d]).to eq :dialog_or_work_sheet
        expect(bs[:dt_d]).to eq(@browser2.wbs[1][:BOF][:dt_d])
        expect(bs[:stName]).to eq 'Worksheet'
      end

      specify 'For dialog sheet "Dialog"' do
        bs = @browser2.globals[:BoundSheet8][1]
        expect(bs[:lbPlyPos]).to eq @browser2.wbs[2][:BOF][:_record][:pos]
        expect(bs[:hsState]).to eq 0
        expect(bs[:hsState_d]).to eq :visible
        expect(bs[:dt]).to eq 0
        expect(bs[:dt_d]).to eq :dialog_or_work_sheet
        expect(bs[:dt_d]).to eq @browser2.wbs[2][:BOF][:dt_d]
        expect(@browser2.wbs[2][:WsBool][:fDialog]).to eq true
        expect(bs[:stName]).to eq 'Dialog'
      end

      # https://superuser.com/questions/1253212/what-is-macro-worksheet-in-excel
      specify 'For macro sheet "Excel 4 macro"' do
        bs = @browser2.globals[:BoundSheet8][2]
        expect(bs[:lbPlyPos]).to eq @browser2.wbs[3][:BOF][:_record][:pos]
        expect(bs[:hsState]).to eq 0
        expect(bs[:hsState_d]).to eq :visible
        expect(bs[:dt]).to eq 1
        expect(bs[:dt_d]).to eq :macro
        expect(bs[:dt_d]).to eq @browser2.wbs[3][:BOF][:dt_d]
        expect(bs[:stName]).to eq 'Excel 4 macro'
      end

      specify 'For macro sheet "International macro"' do
        bs = @browser2.globals[:BoundSheet8][3]
        expect(bs[:lbPlyPos]).to eq @browser2.wbs[4][:BOF][:_record][:pos]
        expect(bs[:hsState]).to eq 0
        expect(bs[:hsState_d]).to eq :visible
        expect(bs[:dt]).to eq 1
        expect(bs[:dt_d]).to eq :macro
        expect(bs[:dt_d]).to eq @browser2.wbs[4][:BOF][:dt_d]
        expect(bs[:stName]).to eq 'International macro'
      end

      specify 'For chart sheet "Chart"' do
        bs = @browser2.globals[:BoundSheet8][4]
        expect(bs[:lbPlyPos]).to eq @browser2.wbs[5][:BOF][:_record][:pos]
        expect(bs[:hsState]).to eq 0
        expect(bs[:hsState_d]).to eq :visible
        expect(bs[:dt]).to eq 2
        expect(bs[:dt_d]).to eq :chart
        expect(bs[:dt_d]).to eq @browser2.wbs[5][:BOF][:dt_d]
        expect(bs[:stName]).to eq 'Chart'
      end

      specify 'For a hidden worksheet' do
        bs = @browser2.globals[:BoundSheet8][5]
        expect(bs[:hsState_d]).to eq :hidden
        expect(bs[:stName]).to eq 'Hidden'
      end

      specify 'For a very hidden worksheet' do
        bs = @browser2.globals[:BoundSheet8][6]
        expect(bs[:hsState_d]).to eq :very_hidden
        expect(bs[:stName]).to eq 'Very hidden'
      end
    end

    context 'DXF record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'tablestyle.xls'))
      end

      let(:dxf) { @browser.get_tse_dxf('Custom Table Style', :whole_table) }

      it 'Decodes basic fields' do
        expect(dxf[:frtHeader]).to be
        expect(dxf[:fNewBorder]).to eq true
        expect(dxf[:cprops]).to eq 14
      end

      it 'Decodes structure of xfPropType 0x00 (FillPattern)' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 0 }
        expect(prop[:cb]).to eq 5
        expect(prop[:_description]).to eq :fill_pattern
        expect(prop[:_structure]).to eq :FillPattern
        expect(prop[:xfPropDataBlob_d]).to eq :FLSGRAY0625
      end

      # Fully tested in StyleExt
      it 'Decodes structures already tested in StyleExt' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 1 } # foreground_color
        expect(prop[:xfPropDataBlob_d][:dwRgba]).to eq :'2A8EF2FF'

        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 2 } # background_color
        expect(prop[:xfPropDataBlob_d][:dwRgba]).to eq :'7030A0FF'

        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 5 } # text_color
        expect(prop[:xfPropDataBlob_d][:dwRgba]).to eq :'2A8EF2FF'

        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 6 } # top_border
        expect(prop[:xfPropDataBlob_d][:dgBorder_d]).to eq :DOTTED

        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 7 } # bottom_border
        expect(prop[:xfPropDataBlob_d][:dgBorder_d]).to eq :DASHDOT

        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 8 } # left_border
        expect(prop[:xfPropDataBlob_d][:dgBorder_d]).to eq :HAIR

        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 9 } # right_border
        expect(prop[:xfPropDataBlob_d][:dgBorder_d]).to eq :DASHDOTDOT
      end

      it 'Decodes structure of xfPropType 0x19 (Bold, font weight)' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 0x19 }
        expect(prop[:cb]).to eq 6
        expect(prop[:_description]).to eq :font_weight
        expect(prop[:_structure]).to eq :Bold
        expect(prop[:xfPropDataBlob_d]).to eq :BLSBOLD
      end

      it 'Decodes structure of xfPropType 0x1A (Underline, underline style)' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 0x1A }
        expect(prop[:cb]).to eq 6
        expect(prop[:_description]).to eq :underline_style
        expect(prop[:_structure]).to eq :Underline
        expect(prop[:xfPropDataBlob_d]).to eq :ULSSINGLE
      end

      it 'Decodes structure of xfPropType 0x1C (_bool1b, text is italicized)' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 0x1C }
        expect(prop[:cb]).to eq 5
        expect(prop[:_description]).to eq :italic
        expect(prop[:_structure]).to eq :_bool1b
        expect(prop[:xfPropDataBlob_d]).to eq true
      end

      it 'Decodes structure of xfPropType 0x1D (_bool1b, text has strikethrough formatting)' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 0x1D }
        expect(prop[:cb]).to eq 5
        expect(prop[:_description]).to eq :strikethrough
        expect(prop[:_structure]).to eq :_bool1b
        expect(prop[:xfPropDataBlob_d]).to eq true
      end

      it 'Decodes structure of xfPropType 0x0B (XFPropBorder, vertical border formatting)' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 0x0B }
        expect(prop[:cb]).to eq 14
        expect(prop[:_description]).to eq :vertical_border
        expect(prop[:_structure]).to eq :XFPropBorder

        prop_data = prop[:xfPropDataBlob_d]
        expect(prop_data[:dgBorder]).to eq 1
        expect(prop_data[:dgBorder_d]).to eq :THIN

        color = prop_data[:color]
        expect(color[:fValidRGBA]).to eq true
        expect(color[:xclrType]).to eq 2
        expect(color[:xclrType_d]).to eq :XCLRRGB
        expect(color[:icv]).to eq 0xFF
        expect(color[:nTintShade]).to eq 0
        expect(color[:nTintShade_d]).to eq 0.0
        expect(color[:dwRgba]).to eq :'2A8EF2FF'
      end

      it 'Decodes structure of xfPropType 0x0C (XFPropBorder, horizontal border formatting)' do
        prop = dxf[:xfPropArray].find { |e| e[:xfPropType] == 0x0C }
        expect(prop[:cb]).to eq 14
        expect(prop[:_description]).to eq :horizontal_border
        expect(prop[:_structure]).to eq :XFPropBorder

        prop_data = prop[:xfPropDataBlob_d]
        expect(prop_data[:dgBorder]).to eq 3
        expect(prop_data[:dgBorder_d]).to eq :DASHED

        color = prop_data[:color]
        expect(color[:fValidRGBA]).to eq true
        expect(color[:xclrType]).to eq 2
        expect(color[:xclrType_d]).to eq :XCLRRGB
        expect(color[:icv]).to eq 0xFF
        expect(color[:nTintShade]).to eq 0
        expect(color[:nTintShade_d]).to eq 0.0
        expect(color[:dwRgba]).to eq :'92D050FF'
      end
    end

    context 'FilePass record' do
      it 'Decodes from files encrypted using XOR obfuscation' do
        file = testfile('biff8', 'filepass', 'filepass_xor-password.xls')
        browser = Unxls::Biff8::Browser.new(file, password: 'password')

        fp = browser.globals[:FilePass]
        expect(fp[:wEncryptionType]).to eq 0
        expect(fp[:_type]).to eq :XOR
        expect(fp[:key]).to eq 5242
        expect(fp[:verificationBytes]).to eq 33711

        expect(browser.globals[:Font][0][:fontName]).to eq 'Arial'
      end

      it 'Decodes from files encrypted using RC4 encryption, RC4 header structure' do
        file = testfile('biff8', 'filepass', 'filepass_rc4_97_2000-password.xls')
        browser = Unxls::Biff8::Browser.new(file, password: 'password')

        fp = browser.globals[:FilePass]
        expect(fp[:wEncryptionType]).to eq 1
        expect(fp[:_type]).to eq :RC4
        expect(fp[:EncryptionVersionInfo]).to eq({ vMajor: 1, vMinor: 1 })
        %i(Salt EncryptedVerifier EncryptedVerifierHash).each do |attr|
          expect(fp[attr]).to be_an_instance_of String
          expect(fp[attr].size).to eq 16
          expect(fp[attr].encoding.name).to eq 'ASCII-8BIT'
        end

        expect(browser.globals[:Font][0][:fontName]).to eq 'Arial'
      end

      [
        { file: 'filepass_cryptoapi_40_basecrypto-password.xls', CSPName: 'Microsoft Base Cryptographic Provider v1.0', KeySize: 40 },
        { file: 'filepass_cryptoapi_40_dhsc-password.xls', CSPName: 'Microsoft DH SChannel Cryptographic Provider', KeySize: 40 },
        { file: 'filepass_cryptoapi_40_dss-password.xls', CSPName: 'Microsoft Base DSS and Diffie-Hellman Cryptographic Provider', KeySize: 40 },
        { file: 'filepass_cryptoapi_128_enhanced-password.xls', CSPName: 'Microsoft Enhanced Cryptographic Provider v1.0', KeySize: 128 },
        { file: 'filepass_cryptoapi_128_enhdss-password.xls', CSPName: 'Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider', KeySize: 128 },
        { file: 'filepass_cryptoapi_128_enhrsa-password.xls', CSPName: 'Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)', KeySize: 128 },
        { file: 'filepass_cryptoapi_128_enhrsasc-password.xls', CSPName: 'Microsoft RSA SChannel Cryptographic Provider', KeySize: 128 },
        { file: 'filepass_cryptoapi_128_strong-password.xls', CSPName: 'Microsoft Strong Cryptographic Provider', KeySize: 128 },
      ].each do |params|
        it "Decodes from #{params[:file]} encrypted using RC4 encryption, RC4 CryptoAPI header structure" do
          file = testfile('biff8', 'filepass', params[:file])
          browser = Unxls::Biff8::Browser.new(file, password: 'password')
          fp = browser.globals[:FilePass]
          expect(fp[:wEncryptionType]).to eq 1
          expect(fp[:_type]).to eq :CryptoAPI
          expect(fp[:EncryptionVersionInfo]).to eq({ vMajor: 2, vMinor: 2 })
          flags = { fCryptoAPI: true, fDocProps: false, fExternal: false, fAES: false }
          expect(fp[:EncryptionHeaderFlags]).to eq flags
          expect(fp[:EncryptionHeaderSize]).to be > 0

          eh = fp[:EncryptionHeader]
          expect(eh[:Flags]).to eq flags
          expect(eh[:SizeExtra]).to eq 0
          expect(eh[:AlgID]).to eq 0x6801
          expect(eh[:AlgID_d]).to eq :RC4
          expect(eh[:KeySize]).to eq params[:KeySize]
          expect(eh[:ProviderType]).to be_an_instance_of Integer
          expect(eh[:CSPName]).to eq params[:CSPName]

          ev = fp[:EncryptionVerifier]
          %i(Salt EncryptedVerifier).each do |attr|
            expect(ev[attr]).to be_an_instance_of String
            expect(ev[attr].size).to eq ev[:SaltSize]
            expect(ev[attr].encoding.name).to eq 'ASCII-8BIT'
          end
          expect(ev[:VerifierHashSize]).to eq 20
          expect(ev[:EncryptedVerifierHash].size).to eq ev[:VerifierHashSize]
          expect(ev[:EncryptedVerifierHash]).to be_an_instance_of String
          expect(ev[:EncryptedVerifierHash].encoding.name).to eq 'ASCII-8BIT'

          expect(fp[:_encryption_algorithm]).to eq :RC4
          expect(fp[:_hashing_algorithm]).to eq :SHA1

          expect(browser.globals[:Font][0][:fontName]).to eq 'Arial'
        end
      end
    end

    context 'Font record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'font.xls'))
      end

      def get_font(stream, row, col)
        font_index = @browser.get_xf(stream, row, col)[:ifnt]
        @browser.get_font(font_index)
      end

      # Size
      it 'Decodes dyHeight field' do
        expect(get_font(1, 2, 0)[:dyHeight]).to eq 200
        expect(get_font(1, 3, 0)[:dyHeight]).to eq 240
        expect(get_font(1, 4, 0)[:dyHeight]).to eq 300
      end

      it 'Decodes fItalic field' do
        expect(get_font(1, 2, 1)[:fItalic]).to eq true
      end

      it 'Decodes fStrikeOut field' do
        expect(get_font(1, 2, 2)[:fStrikeOut]).to eq true
      end

      # Seems like fOutline, fShadow are only used in old (~2) versions of Excel for Mac
      # Seems like fCondense, fExtend came from Word but can't be found in the interface of Excel

      # Color
      it 'Decodes icv field' do
        expect(get_font(1, 2, 3)[:icv]).to eq 0x7FFF # Font automatic color
        expect(get_font(1, 3, 3)[:icv]).to eq 0x000A # rgColor[2] of Palette, FF0000
      end

      # Bold
      it 'Decodes bls field' do
        f = get_font(1, 2, 4)
        expect(f[:bls]).to eq 400
        expect(f[:bls_d]).to eq :BLSNORMAL

        f = get_font(1, 3, 4)
        expect(f[:bls]).to eq 700
        expect(f[:bls_d]).to eq :BLSBOLD
      end

      # Superscript, subscript
      it 'Decodes sss field' do
        f = get_font(1, 2, 5)
        expect(f[:sss]).to eq 0
        expect(f[:sss_d]).to eq :SSSNONE

        f = get_font(1, 3, 5)
        expect(f[:sss]).to eq 1
        expect(f[:sss_d]).to eq :SSSSUPER

        f = get_font(1, 4, 5)
        expect(f[:sss]).to eq 2
        expect(f[:sss_d]).to eq :SSSSUB
      end

      # Underline
      it 'Decodes uls field' do
        f = get_font(1, 2, 6)
        expect(f[:uls]).to eq 0
        expect(f[:uls_d]).to eq :ULSNONE

        f = get_font(1, 3, 6)
        expect(f[:uls]).to eq 1
        expect(f[:uls_d]).to eq :ULSSINGLE

        f = get_font(1, 4, 6)
        expect(f[:uls]).to eq 2
        expect(f[:uls_d]).to eq :ULSDOUBLE

        f = get_font(1, 5, 6)
        expect(f[:uls]).to eq 33
        expect(f[:uls_d]).to eq :ULSSINGLEACCOUNTANT

        f = get_font(1, 6, 6)
        expect(f[:uls]).to eq 34
        expect(f[:uls_d]).to eq :ULSDOUBLEACCOUNTANT
      end

      # Font family (Mac versions seem to not to write this field)
      it 'Decodes bFamily field' do
        f = get_font(1, 2, 7)
        expect(f[:bFamily]).to eq 0
        expect(f[:bFamily_d]).to eq :'Not applicable'

        f = get_font(1, 3, 7)
        expect(f[:bFamily]).to eq 1
        expect(f[:bFamily_d]).to eq :Roman

        f = get_font(1, 4, 7)
        expect(f[:bFamily]).to eq 2
        expect(f[:bFamily_d]).to eq :Swiss

        f = get_font(1, 5, 7)
        expect(f[:bFamily]).to eq 3
        expect(f[:bFamily_d]).to eq :Modern

        f = get_font(1, 6, 7)
        expect(f[:bFamily]).to eq 4
        expect(f[:bFamily_d]).to eq :Script
      end

      # Font charset (Mac versions seem to not to write this field)
      it 'Decodes bCharSet field' do
        f = get_font(1, 2, 8)
        expect(f[:bCharSet]).to eq 128
        expect(f[:bCharSet_d]).to eq :ShiftJIS

        f = get_font(1, 3, 8)
        expect(f[:bCharSet]).to eq 129
        expect(f[:bCharSet_d]).to eq :Jangul

        f = get_font(1, 4, 8)
        expect(f[:bCharSet]).to eq 136
        expect(f[:bCharSet_d]).to eq :ChineseBIG5

        f = get_font(1, 5, 8)
        expect(f[:bCharSet]).to eq 134
        expect(f[:bCharSet_d]).to eq :GB2312

        f = get_font(1, 6, 8)
        expect(f[:bCharSet]).to eq 0
        expect(f[:bCharSet_d]).to eq :ANSI
      end

      it 'Decodes fontName field' do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'font-fontname.xls'))

        expect(get_font(1, 2, 0)[:fontName]).to eq 'HGP創英角ｺﾞｼｯｸUB'
        expect(get_font(1, 3, 0)[:fontName]).to eq 'メイリオ'
        expect(get_font(1, 4, 0)[:fontName]).to eq '나눔고딕 ExtraBold'
        expect(get_font(1, 5, 0)[:fontName]).to eq 'مِصحفي'
        expect(get_font(1, 6, 0)[:fontName]).to eq 'Цветные эмодзи Apple'
        expect(get_font(1, 7, 0)[:fontName]).to eq 'Arial'
      end
    end

    context 'Format record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'format.xls'))
      end

      # Record id
      it 'Decodes ifmt field' do
        1.upto(5).each do |row|
          expect(@browser.get_format(1, row, 0)[:ifmt]).to be_an_instance_of Integer
        end
      end

      # Format string
      it 'Decodes stFormat field' do
        expect(@browser.get_format(1, 1, 0)[:stFormat]).to eq '"Format フォーマット"'
        expect(@browser.get_format(1, 2, 0)[:stFormat]).to eq '[$-409]d\-mmm\-yy;@'
        expect(@browser.get_format(1, 3, 0)[:stFormat]).to eq 'm"月"d"日";@'
        expect(@browser.get_format(1, 4, 0)[:stFormat]).to eq 'dd/mm/yy;@'
        expect(@browser.get_format(1, 5, 0)[:stFormat]).to eq '_-[$₩-412]* #,##0.00_-;\-[$₩-412]* #,##0.00_-;_-[$₩-412]* "-"??_-;_-@_-'
      end
    end

    context 'Palette record' do
      before(:context) do
        @browser1 = Unxls::Biff8::Browser.new(testfile('biff8', 'empty.xls'))
        @browser2 = Unxls::Biff8::Browser.new(testfile('biff8', 'palette.xls'))
      end

      it 'Decodes' do
        expect(@browser1.globals[:Palette]).to eq nil # Uses standard palette

        p = @browser2.globals[:Palette]
        expect(p[:ccv]).to eq 56
        expect(p[:rgColor].size).to eq p[:ccv]
        expect(p[:rgColor][0]).to eq :"2A8EF200" # Black changed to RGB 42, 142, 242
        expect(p[:rgColor][1]).to eq :"F28E2A00" # White changed to RGB 242, 142, 142
        expect(p[:rgColor][2]).to eq :"8E2AF200" # Red changed to RGB 142, 42, 242
        expect(p[:rgColor][3]).to eq :"00FF0000" # Green unchanged
      end
    end

    context 'SST record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'sst.xls'))
      end

      it 'Decodes cstTotal field' do
        expect(@browser.globals[:SST][:cstTotal]).to eq 24
      end

      it 'Decodes cstUnique field' do
        expect(@browser.globals[:SST][:cstUnique]).to eq 20
      end

      describe 'String-related fields' do
        specify 'For short ASCII string' do
          sst = @browser.get_sst(1, 3, 0)
          expect(sst[:fHighByte]).to eq false
          expect(sst[:fRichSt]).to eq false
          expect(sst[:cch]).to eq 35
          expect(sst[:rgb]).to eq 'String with single-byte characters.'
          expect(sst[:rgb].size).to eq sst[:cch]
        end

        specify 'For short string with single and double-byte characters' do
          sst = @browser.get_sst(1, 4, 0)
          expect(sst[:fHighByte]).to eq true
          expect(sst[:fRichSt]).to eq false
          expect(sst[:cch]).to eq 49
          expect(sst[:rgb]).to eq 'String with double-byte characters: АБВГД, アイウエオ。'
          expect(sst[:rgb].size).to eq sst[:cch]
        end

        specify 'For very long string with surrogate pair characters' do
          sst = @browser.get_sst(1, 7, 0)
          expect(sst[:fHighByte]).to eq true
          expect(sst[:fRichSt]).to eq false
          expect(sst[:rgb].size).to eq 17000
          expect(sst[:cch] - sst[:rgb].size).to eq 5 # 1 surrogate pair character (emojis, etc.) is counted as 2
          expect(sst[:rgb][0..9]).to eq 'Very long '
          expect(sst[:rgb][-9..-1]).to eq ' W W END.'
        end
      end

      describe 'Formatting runs-related fields' do
        let(:sst) do
          @browser.get_sst(1, 3, 1)
        end

        it 'Decodes basic fields' do
          expect(sst[:fRichSt]).to eq true
          expect(sst[:cRun]).to eq 10
          expect(sst[:rgb][0..10]).to eq 'Lorem ipsum'
        end

        it 'Decodes first run ("ipsum")' do
          expect(sst[:rgRun][0][:ich]).to eq 6 # Run's start character position
          expect(sst[:rgRun][1][:ich]).to eq 11 # Run's end character position + 1
          frun_ifnt = sst[:rgRun][0][:ifnt] - 1 # See 2.5.129 FontIndex
          font = @browser.globals[:Font][frun_ifnt]
          expect(font[:fontName]).to eq 'Meiryo'
          expect(sst[:rgRun][1][:ifnt]).to eq 0
        end

        it 'Decodes second run ("amet")' do
          expect(sst[:rgRun][2][:ich]).to eq 22
          expect(sst[:rgRun][3][:ich]).to eq 26
          frun_ifnt = sst[:rgRun][2][:ifnt] - 1
          font = @browser.globals[:Font][frun_ifnt]
          expect(font[:fontName]).to eq 'Osaka'
          expect(sst[:rgRun][3][:ifnt]).to eq 0
        end

        it 'Decodes third run ("adipiscing")' do
          expect(sst[:rgRun][4][:ich]).to eq 40
          expect(sst[:rgRun][5][:ich]).to eq 50
          frun_ifnt = sst[:rgRun][4][:ifnt] - 1
          font = @browser.globals[:Font][frun_ifnt]
          expect(font[:fontName]).to eq 'FangSong'
          expect(sst[:rgRun][5][:ifnt]).to eq 0
        end

        # etc
      end

      context 'Phonetic text runs-related fields' do
        it 'Decodes basic fields' do
          sst = @browser.get_sst(1, 3, 2)
          expect(sst[:fExtSt]).to eq true
          expect(sst[:cbExtRst]).to eq 92
          expect(sst[:rgb]).to eq '世界選手権でメダルを取った選手が池江選手にメッセージ'
          expect(sst[:ExtRst][:cb]).to eq 88
        end

        it 'Decodes rphssub field' do
          rphssub = @browser.get_sst(1, 3, 2)[:ExtRst][:rphssub]
          expect(rphssub[:crun]).to eq 6
          expect(rphssub[:cch]).to eq 21
          expect(rphssub[:cch]).to eq rphssub[:st].size
          expect(rphssub[:st]).to eq 'ｾｶｲｾﾝｼｭｹﾝﾄｾﾝｼｭｲｹｴｾﾝｼｭ'
        end

        it 'Decodes rgphruns field' do
          rgphruns = @browser.get_sst(1, 3, 2)[:ExtRst][:rgphruns]

          # First run: "ｾｶｲ" (0-) over "世界" (0-1)
          expect(rgphruns[0][:ichFirst]).to eq 0 # furigana (rphssub.st) start char
          expect(rgphruns[0][:ichMom]).to eq 0 # text (SST.rgb[].rgb) start char
          expect(rgphruns[0][:cchMom]).to eq 2 # text char count

          # Second run: "ｾﾝｼｭｹﾝ" (3-) over "選手権" (2-4)
          expect(rgphruns[1][:ichFirst]).to eq 3
          expect(rgphruns[1][:ichMom]).to eq 2
          expect(rgphruns[1][:cchMom]).to eq 3

          # Third run: "ﾄ" (9-) over "取" (10)
          expect(rgphruns[2][:ichFirst]).to eq 9
          expect(rgphruns[2][:ichMom]).to eq 10
          expect(rgphruns[2][:cchMom]).to eq 1

          # 4th run: "ｾﾝｼｭ" (10-) over "選手" (14-15)
          expect(rgphruns[3][:ichFirst]).to eq 10
          expect(rgphruns[3][:ichMom]).to eq 13
          expect(rgphruns[3][:cchMom]).to eq 2

          # etc
        end

        it 'Decodes phs field' do
          phs = @browser.get_sst(1, 3, 2)[:ExtRst][:phs]
          expect(phs[:phType]).to eq 0
          expect(phs[:phType_d]).to eq :narrow_katakana

          phs = @browser.get_sst(1, 4, 2)[:ExtRst][:phs]
          expect(phs[:phType]).to eq 1
          expect(phs[:phType_d]).to eq :wide_katakana

          phs = @browser.get_sst(1, 5, 2)[:ExtRst][:phs]
          expect(phs[:phType]).to eq 2
          expect(phs[:phType_d]).to eq :hiragana

          phs = @browser.get_sst(1, 3, 3)[:ExtRst][:phs]
          expect(phs[:alcH]).to eq 0
          expect(phs[:alcH_d]).to eq :general

          phs = @browser.get_sst(1, 4, 3)[:ExtRst][:phs]
          expect(phs[:alcH]).to eq 1
          expect(phs[:alcH_d]).to eq :left

          phs = @browser.get_sst(1, 5, 3)[:ExtRst][:phs]
          expect(phs[:alcH]).to eq 2
          expect(phs[:alcH_d]).to eq :center

          phs = @browser.get_sst(1, 6, 3)[:ExtRst][:phs]
          expect(phs[:alcH]).to eq 3
          expect(phs[:alcH_d]).to eq :distributed

          phs = @browser.get_sst(1, 3, 4)[:ExtRst][:phs]
          phs_ifnt = phs[:ifnt] - 1
          font = @browser.globals[:Font][phs_ifnt]
          expect(font[:fontName]).to eq 'HGPMinchoE'

          phs = @browser.get_sst(1, 4, 4)[:ExtRst][:phs]
          phs_ifnt = phs[:ifnt] - 1
          font = @browser.globals[:Font][phs_ifnt]
          expect(font[:fontName]).to eq 'Hiragino Kaku Gothic StdN W8'
        end
      end
    end

    context 'Style record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'style-styleext.xls'))
      end

      it 'Decodes fields for built-in styles' do
        s = @browser.get_style(1, 2, 0)
        expect(s[:ixfe]).to eq 0
        expect(s[:fBuiltIn]).to eq true
        expect(s[:builtInData][:istyBuiltIn]).to eq 0
        expect(s[:builtInData][:iLevel]).to eq 0xFF
        expect(s[:builtInData][:istyBuiltIn_d]).to eq :Normal

        s = @browser.get_style(1, 3, 0)
        expect(s[:ixfe]).to eq 43
        expect(s[:fBuiltIn]).to eq true
        expect(s[:builtInData][:istyBuiltIn]).to eq 3
        expect(s[:builtInData][:iLevel]).to eq 0xFF
        expect(s[:builtInData][:istyBuiltIn_d]).to eq :Comma
      end

      it 'Decodes fields for basic styles written by Excel' do
        s = @browser.get_style(1, 2, 1)
        expect(s[:ixfe]).to eq 40
        expect(s[:fBuiltIn]).to eq false
        expect(s[:user]).to eq :Bad

        s = @browser.get_style(1, 3, 1)
        expect(s[:ixfe]).to eq 41
        expect(s[:fBuiltIn]).to eq false
        expect(s[:user]).to eq :Calculation
      end

      it 'Decodes custom styles' do
        s = @browser.get_style(1, 4, 2)
        expect(s[:ixfe]).to eq 47
        expect(s[:fBuiltIn]).to eq false
        expect(s[:user]).to eq :'Custom 1'
      end
    end

    context 'StyleExt record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'style-styleext.xls'))
      end

      it 'Decodes basic fields' do
        s = @browser.get_styleext(1, 2, 2) # Modified built-in style
        expect(s[:fBuiltIn]).to eq true
        expect(s[:fHidden]).to eq false
        expect(s[:fCustom]).to eq true
        expect(s[:iCategory]).to eq 5
        expect(s[:iCategory_d]).to eq :'Number format style'
        expect(s[:builtInData]).to eq @browser.get_style(1, 2, 2)[:builtInData]
        expect(s[:stName]).to eq :Currency
        expect(s[:cprops]).to eq 3

        s = @browser.get_styleext(1, 3, 2) # Modified basic style
        expect(s[:fBuiltIn]).to eq true # fBuiltIn in corresponding Style is false
        expect(s[:fHidden]).to eq false
        expect(s[:fCustom]).to eq true
        expect(s[:iCategory]).to eq 3
        expect(s[:iCategory_d]).to eq :'Title and heading style'
        expect(s[:stName]).to eq :Total
        expect(s[:cprops]).to eq 3

        s = @browser.get_styleext(1, 4, 2) # Custom style
        expect(s[:fBuiltIn]).to eq false
        expect(s[:fHidden]).to eq false
        expect(s[:fCustom]).to eq false
        expect(s[:iCategory]).to eq 0
        expect(s[:iCategory_d]).to eq :'Custom style'
        expect(s[:stName]).to eq :'Custom 1'
        expect(s[:cprops]).to eq 10
      end

      context 'Decodes xfPropArray field' do
        let(:custom1) { @browser.get_styleext(1, 4, 2)[:xfPropArray] }
        let(:custom2) { @browser.get_styleext(1, 5, 2)[:xfPropArray] }

        it 'Decodes structure of xfPropType 0x01 (XFPropColor, foreground color)' do
          prop = custom1.find { |e| e[:xfPropType] == 1 }
          expect(prop[:cb]).to eq 12
          expect(prop[:_description]).to eq :foreground_color
          expect(prop[:_structure]).to eq :XFPropColor

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:fValidRGBA]).to eq true
          expect(prop_data[:xclrType]).to eq 3
          expect(prop_data[:xclrType_d]).to eq :XCLRTHEMED
          expect(prop_data[:icv]).to eq 5
          expect(prop_data[:nTintShade]).to eq 0
          expect(prop_data[:nTintShade_d]).to eq 0.0
          expect(prop_data[:dwRgba]).to eq :ED7D31FF
        end

        it 'Decodes structure of xfPropType 0x02 (XFPropColor, background color)' do
          prop = custom1.find { |e| e[:xfPropType] == 2 }
          expect(prop[:cb]).to eq 12
          expect(prop[:_description]).to eq :background_color
          expect(prop[:_structure]).to eq :XFPropColor

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:fValidRGBA]).to eq true
          expect(prop_data[:xclrType]).to eq 2
          expect(prop_data[:xclrType_d]).to eq :XCLRRGB
          expect(prop_data[:icv]).to eq 0xFF
          expect(prop_data[:nTintShade]).to eq 0
          expect(prop_data[:nTintShade_d]).to eq 0.0
          expect(prop_data[:dwRgba]).to eq :'2A8EF2FF' # rgb(42, 142, 242)
        end

        it 'Decodes structure of xfPropType 0x03 (XFPropGradient, gradient fill)' do
          prop = custom2.find { |e| e[:xfPropType] == 3 }
          expect(prop[:cb]).to eq 48
          expect(prop[:_description]).to eq :gradient_fill
          expect(prop[:_structure]).to eq :XFPropGradient

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:type]).to eq 0
          expect(prop_data[:type_d]).to eq :linear
          expect(prop_data[:numDegree]).to eq 45.0
          expect(prop_data[:numFillToLeft]).to eq 0.0
          expect(prop_data[:numFillToRight]).to eq 0.0
          expect(prop_data[:numFillToTop]).to eq 0.0
          expect(prop_data[:numFillToBottom]).to eq 0.0
        end

        it 'Decodes structure of xfPropType 0x04 (XFPropGradientStop, gradient stop)' do
          props = custom2.select { |e| e[:xfPropType] == 4 }

          {
            0 => 0.0,
            2 => 1.0
          }.each do |prop_index, num_position|
            prop = props[prop_index]
            expect(prop[:cb]).to eq 22
            expect(prop[:_description]).to eq :gradient_stop
            expect(prop[:_structure]).to eq :XFPropGradientStop

            prop_data = prop[:xfPropDataBlob_d]
            expect(prop_data[:numPosition]).to eq num_position

            color = prop_data[:color]
            expect(color[:fValidRGBA]).to eq true
            expect(color[:xclrType]).to eq 3
            expect(color[:xclrType_d]).to eq :XCLRTHEMED
            expect(color[:icv]).to eq 7
            expect(color[:nTintShade]).to eq -8224
            expect(color[:nTintShade_d]).to eq -0.25
            expect(color[:dwRgba]).to eq :BD8E00FF
          end

          prop = props[1]
          expect(prop[:cb]).to eq 22
          expect(prop[:_description]).to eq :gradient_stop
          expect(prop[:_structure]).to eq :XFPropGradientStop

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:numPosition]).to eq 0.5

          color = prop_data[:color]
          expect(color[:fValidRGBA]).to eq true
          expect(color[:xclrType]).to eq 3
          expect(color[:xclrType_d]).to eq :XCLRTHEMED
          expect(color[:icv]).to eq 9
          expect(color[:nTintShade]).to eq 13107
          expect(color[:nTintShade_d]).to eq 0.4
          expect(color[:dwRgba]).to eq :A9D08EFF
        end

        it 'Decodes structure of xfPropType 0x05 (XFPropColor, text color)' do
          prop = custom1.find { |e| e[:xfPropType] == 5 }
          expect(prop[:cb]).to eq 12
          expect(prop[:_description]).to eq :text_color
          expect(prop[:_structure]).to eq :XFPropColor

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:fValidRGBA]).to eq true
          expect(prop_data[:xclrType]).to eq 3
          expect(prop_data[:xclrType_d]).to eq :XCLRTHEMED
          expect(prop_data[:icv]).to eq 5
          expect(prop_data[:nTintShade]).to eq -8190
          expect(prop_data[:nTintShade_d]).to eq -0.25
          expect(prop_data[:dwRgba]).to eq :C65911FF
        end

        it 'Decodes structure of xfPropType 0x06 (XFPropBorder, top border formatting)' do
          prop = custom1.find { |e| e[:xfPropType] == 6 }
          expect(prop[:cb]).to eq 14
          expect(prop[:_description]).to eq :top_border
          expect(prop[:_structure]).to eq :XFPropBorder

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:dgBorder]).to eq 8
          expect(prop_data[:dgBorder_d]).to eq :MEDIUMDASHED

          color = prop_data[:color]
          expect(color[:fValidRGBA]).to eq true
          expect(color[:xclrType]).to eq 3
          expect(color[:xclrType_d]).to eq :XCLRTHEMED
          expect(color[:icv]).to eq 9
          expect(color[:nTintShade]).to eq 0
          expect(color[:nTintShade_d]).to eq 0.0
          expect(color[:dwRgba]).to eq :'70AD47FF'
        end

        it 'Decodes structure of xfPropType 0x07 (XFPropBorder, bottom border formatting)' do
          prop = custom1.find { |e| e[:xfPropType] == 7 }
          expect(prop[:cb]).to eq 14
          expect(prop[:_description]).to eq :bottom_border
          expect(prop[:_structure]).to eq :XFPropBorder

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:dgBorder]).to eq 1
          expect(prop_data[:dgBorder_d]).to eq :THIN

          color = prop_data[:color]
          expect(color[:fValidRGBA]).to eq true
          expect(color[:xclrType]).to eq 2
          expect(color[:xclrType_d]).to eq :XCLRRGB
          expect(color[:icv]).to eq 0xFF
          expect(color[:nTintShade]).to eq 0
          expect(color[:nTintShade_d]).to eq 0.0
          expect(color[:dwRgba]).to eq :'7030A0FF'
        end

        it 'Decodes structure of xfPropType 0x08 (XFPropBorder, left border formatting)' do
          prop = custom1.find { |e| e[:xfPropType] == 8 }
          expect(prop[:cb]).to eq 14
          expect(prop[:_description]).to eq :left_border
          expect(prop[:_structure]).to eq :XFPropBorder

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:dgBorder]).to eq 5
          expect(prop_data[:dgBorder_d]).to eq :THICK

          color = prop_data[:color]
          expect(color[:fValidRGBA]).to eq true
          expect(color[:xclrType]).to eq 2
          expect(color[:xclrType_d]).to eq :XCLRRGB
          expect(color[:icv]).to eq 0xFF
          expect(color[:nTintShade]).to eq 0
          expect(color[:nTintShade_d]).to eq 0.0
          expect(color[:dwRgba]).to eq :'2A8EF2FF'
        end

        it 'Decodes structure of xfPropType 0x09 (XFPropBorder, right border formatting)' do
          prop = custom1.find { |e| e[:xfPropType] == 9 }
          expect(prop[:cb]).to eq 14
          expect(prop[:_description]).to eq :right_border
          expect(prop[:_structure]).to eq :XFPropBorder

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:dgBorder]).to eq 3
          expect(prop_data[:dgBorder_d]).to eq :DASHED

          color = prop_data[:color]
          expect(color[:fValidRGBA]).to eq true
          expect(color[:xclrType]).to eq 2
          expect(color[:xclrType_d]).to eq :XCLRRGB
          expect(color[:icv]).to eq 0xFF
          expect(color[:nTintShade]).to eq 0
          expect(color[:nTintShade_d]).to eq 0.0
          expect(color[:dwRgba]).to eq :FF0000FF
        end

        it 'Decodes structure of xfPropType 0x0A (XFPropBorder, diagonal border formatting)' do
          prop = custom1.find { |e| e[:xfPropType] == 10 }
          expect(prop[:cb]).to eq 14
          expect(prop[:_description]).to eq :diagonal_border
          expect(prop[:_structure]).to eq :XFPropBorder

          prop_data = prop[:xfPropDataBlob_d]
          expect(prop_data[:dgBorder]).to eq 9
          expect(prop_data[:dgBorder_d]).to eq :DASHDOT

          color = prop_data[:color]
          expect(color[:fValidRGBA]).to eq true
          expect(color[:xclrType]).to eq 3
          expect(color[:xclrType_d]).to eq :XCLRTHEMED
          expect(color[:icv]).to eq 7
          expect(color[:nTintShade]).to eq 0
          expect(color[:nTintShade_d]).to eq 0.0
          expect(color[:dwRgba]).to eq :FFC000FF
        end

        it 'Decodes structure of 0x0D (diagonal up border)' do
          prop = custom1.find { |e| e[:xfPropType] == 13 }
          expect(prop[:cb]).to eq 5
          expect(prop[:_description]).to eq :diagonal_up
          expect(prop[:_structure]).to eq :_bool1b
          expect(prop[:xfPropDataBlob_d]).to eq true
        end

        it 'Decodes structure of 0x0E (diagonal down)' do
          prop = custom1.find { |e| e[:xfPropType] == 14 }
          expect(prop[:cb]).to eq 5
          expect(prop[:_description]).to eq :diagonal_down
          expect(prop[:_structure]).to eq :_bool1b
          expect(prop[:xfPropDataBlob_d]).to eq true
        end
      end

      # It seems that in StyleExt record Excel uses only the xfPropTypes that specify
      # color, to render the resuling RGB values (e.g. themed color with applied tint,
      # or RGB color) to the dwRgba field. Other properties, even though they are
      # specified for the style, are written to style XFs. At any rate, there seems
      # no way to make Excel to write out these properties to a StyleExt record using
      # the UI.
      context 'Seemingly unused xfPropTypes' do
        # it 'Decodes structure of xfPropType 0x00 (FillPattern)' # used by TableStyleElement/DXF
        # it 'Decodes structure of xfPropType 0x0B (XFPropBorder, vertical border formatting)' # used by TableStyleElement/DXF
        # it 'Decodes structure of xfPropType 0x0C (XFPropBorder, horizontal border formatting)' # used by TableStyleElement/DXF
        it 'Decodes structure of 0x0F (HorizAlign, horizontal alignment)'
        it 'Decodes structure of 0x10 (VertAlign, vertical alignment)'
        it 'Decodes structure of 0x11 (XFPropTextRotation, text rotation)'
        it 'Decodes structure of 0x12 (indentation level)'
        it 'Decodes structure of 0x13 (ReadingOrder, reading order)'
        it 'Decodes structure of 0x14 (text wrap)'
        it 'Decodes structure of 0x15 (justify distributed)'
        it 'Decodes structure of 0x16 (shrink to fit)'
        it 'Decodes structure of 0x17 (cell is merged)'
        it 'Decodes structure of 0x18 (LPWideString, font name)'
        # it 'Decodes structure of 0x19 (Bold, font_weight)' # used by TableStyleElement/DXF
        # it 'Decodes structure of 0x1A (Underline, underline style)' # used by TableStyleElement/DXF
        it 'Decodes structure of 0x1B (Script, script style)'
        # it 'Decodes structure of 0x1C (text is italicized)' # used by TableStyleElement/DXF
        # it 'Decodes structure of 0x1D (text has strikethrough formatting)' # used by TableStyleElement/DXF
        it 'Decodes structure of 0x1E (text has an outline style)'
        it 'Decodes structure of 0x1F (text has a shadow style)'
        it 'Decodes structure of 0x20 (text is condensed)'
        it 'Decodes structure of 0x21 (text is extended)'
        it 'Decodes structure of 0x22 (font character set)'
        it 'Decodes structure of 0x23 (font family)'
        it 'Decodes structure of 0x24 (text size)'
        it 'Decodes structure of 0x25 (FontScheme, font scheme)'
        it 'Decodes structure of 0x26 (XLUnicodeString, number format string)'
        it 'Decodes structure of 0x29 (IFmt, number format identifier)'
        it 'Decodes structure of 0x2A (text relative indentation)'
        it 'Decodes structure of 0x2B (locked)'
        it 'Decodes structure of 0x2C (hidden)'
      end
    end

    context 'TableStyles, TableStyle, TableStyleElement records' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'tablestyle.xls'))
      end

      let(:ts_name) { 'Custom Table Style' }
      let(:pts_name) { 'Custom PivotTable Style' }

      it 'Decodes TableStyles record' do
        record = @browser.globals[:TableStyles]
        expect(record[:frtHeader]).to be
        expect(record[:cts]).to eq 146
        expect(record[:cchDefTableStyle]).to eq 18
        expect(record[:cchDefPivotStyle]).to eq 23
        expect(record[:rgchDefTableStyle]).to eq ts_name
        expect(record[:rgchDefPivotStyle]).to eq pts_name
      end

      it 'Decodes TableStyle record' do
        record = @browser.get_tablestyle(ts_name)
        expect(record[:rgchName]).to eq ts_name
        expect(record[:fIsPivot]).to eq false
        expect(record[:fIsTable]).to eq true
        expect(record[:ctse]).to be_between(1, 28)
        expect(record[:cchName]).to eq 18

        record = @browser.get_tablestyle(pts_name)
        expect(record[:rgchName]).to eq pts_name
        expect(record[:fIsPivot]).to eq true
        expect(record[:fIsTable]).to eq false
        expect(record[:ctse]).to be_between(1, 28)
        expect(record[:cchName]).to eq 23
      end

      context 'TableStyleElement record' do
        [
          { tseType: 0x00, tseType_d: :whole_table }, # Whole Table
          { tseType: 0x01, tseType_d: :header_row }, # Header Row
          { tseType: 0x02, tseType_d: :total_row }, # Grand Total Row
          { tseType: 0x03, tseType_d: :first_column }, # First Column
          { tseType: 0x04, tseType_d: :last_column }, # Grand Total Column
          { tseType: 0x05, tseType_d: :row_stripe_1, size: 4 }, # First Row Stripe
          { tseType: 0x06, tseType_d: :row_stripe_2, size: 5 }, # Second Row Stripe
          { tseType: 0x07, tseType_d: :column_stripe_1, size: 2 }, # First Column Stripe
          { tseType: 0x08, tseType_d: :column_stripe_2, size: 3 }, # Second Column Stripe
          { tseType: 0x09, tseType_d: :first_cell_header }, # First Header Cell
          { tseType: 0x0A, tseType_d: :last_cell_header }, # Last Header Cell
          { tseType: 0x0B, tseType_d: :first_cell_total }, # First Total Cell
          { tseType: 0x0C, tseType_d: :last_cell_total }, # Last Total Cell
          { tseType: 0x0D, tseType_d: :pt_outermost_subtotal_columns }, # Subtotal Column 1
          { tseType: 0x0E, tseType_d: :pt_alternating_even_subtotal_columns }, # Subtotal Column 2
          { tseType: 0x0F, tseType_d: :pt_alternating_odd_subtotal_columns }, # Subtotal Column 3
          { tseType: 0x10, tseType_d: :pt_outermost_subtotal_rows }, # Subtotal Row 1
          { tseType: 0x11, tseType_d: :pt_alternating_even_subtotal_rows }, # Subtotal Row 2
          { tseType: 0x12, tseType_d: :pt_alternating_odd_subtotal_rows }, # Subtotal Row 3
          { tseType: 0x13, tseType_d: :pt_empty_rows_after_each_subtotal_row }, # Blank Row
          { tseType: 0x14, tseType_d: :pt_outermost_column_subheadings }, # Column Subheading 1
          { tseType: 0x15, tseType_d: :pt_alternating_even_column_subheadings }, # Column Subheading 2
          { tseType: 0x16, tseType_d: :pt_alternating_odd_column_subheadings }, # Column Subheading 3
          { tseType: 0x17, tseType_d: :pt_outermost_row_subheadings }, # Row Subheading 1
          { tseType: 0x18, tseType_d: :pt_alternating_even_row_subheadings }, # Row Subheading 2
          { tseType: 0x19, tseType_d: :pt_alternating_odd_row_subheadings }, # Row Subheading 3
          { tseType: 0x1A, tseType_d: :pt_page_field_captions }, # Report Filter Labels
          { tseType: 0x1B, tseType_d: :pt_page_item_captions }, # Report Filter Values
        ].each do |params|
          it "Decodes records of type #{params[:tseType_d]}" do
            style_name = params[:tseType_d].to_s.start_with?('pt_') ? pts_name : ts_name
            element = @browser.get_tablestyleelement(style_name, params[:tseType_d])
            expect(element[:tseType]).to eq params[:tseType]
            expect(element[:tseType_d]).to eq params[:tseType_d]
            if params[:size]
              expect(element[:size]).to eq params[:size]
            else
              expect(element[:size]).to be_between(1, 9)
            end
            expect(element[:index]).to be_an_instance_of Integer
          end
        end

        it 'Adds custom :_tsi field to reference parent TableStyle record' do
          styles = @browser.globals[:TableStyle]
          styles.each do |style|
            index = style[:_record][:index]
            number_of_elements = @browser.globals[:TableStyleElement].select { |r| r[:_tsi] == index }.size
            expect(style[:ctse]).to eq number_of_elements
          end
        end
      end
    end

    context 'XF record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'xf.xls'))
      end

      # Relation to Font record
      it 'Decodes ifnt field' do
        ifnt = @browser.get_xf(1, 2, 0)[:ifnt]
        ifnt -= 1 if ifnt >= 4 # See p. 677
        expect(@browser.globals[:Font][ifnt][:fontName]).to eq 'Mongolian Baiti'
      end

      # Relation to Format record
      it 'Decodes ifmt field' do
        expect(@browser.get_xf(1, 2, 1)[:ifmt]).to eq 0

        ifmt = @browser.get_xf(1, 3, 1)[:ifmt]
        format = @browser.globals[:Format].find { |r| r[:ifmt] == ifmt }
        expect(format[:stFormat]).to eq '"Format フォーマット"'
      end

      # Format - Cells - Protection - Locked
      it 'Decodes fLocked field' do
        expect(@browser.get_xf(1, 2, 2)[:fLocked]).to eq true
        expect(@browser.get_xf(1, 3, 2)[:fLocked]).to eq false
      end

      # Format - Cells - Protection - Hidden
      it 'Decodes fHidden field' do
        expect(@browser.get_xf(1, 2, 3)[:fHidden]).to eq true
        expect(@browser.get_xf(1, 3, 3)[:fHidden]).to eq false
      end

      # Lotus 1-2-3 prefixes for text alignment. Enable (Win-only): Tools - Options - Transition - Transition navigation keys
      # Mac version seems to keep these fields between saves
      it 'Decodes f123Prefix field' do
        expect(@browser.get_xf(1, 2, 4)[:f123Prefix]).to eq false # none
        expect(@browser.get_xf(1, 3, 4)[:f123Prefix]).to eq true # ' single quote 0x27 (left-aligned)
        expect(@browser.get_xf(1, 4, 4)[:f123Prefix]).to eq true # " double quote 0x22 (right-aligned)
        expect(@browser.get_xf(1, 5, 4)[:f123Prefix]).to eq true # ^ caret 0x5E (centered)
        expect(@browser.get_xf(1, 6, 4)[:f123Prefix]).to eq true # \ backslash 0x5C (justified)
      end

      # Specifies whether XF is a style
      it 'Decodes fStyle field' do
        14.times do |i|
          expect(@browser.globals[:XF][i][:fStyle]).to eq true # Built-in style XFs
        end

        expect(@browser.globals[:XF][15][:fStyle]).to eq false # Default cell format XF

        user_style1_xf_index = @browser.get_xf(1, 5, 5)[:ixfParent] # "User style 1" XF
        expect(@browser.globals[:XF][user_style1_xf_index][:fStyle]).to eq true
      end

      # Relation of cell XF to its style XF
      it 'Decodes ixfParent field' do
        expect(@browser.globals[:XF][0][:ixfParent]).to eq 0xFFF # Builtin Normal style XF
        expect(@browser.globals[:XF][15][:ixfParent]).to eq 0 # Default cell format XF
        expect(@browser.get_xf(1, 2, 5)[:ixfParent]).to eq 0 # Cell set to Normal style
        expect(@browser.get_xf(1, 3, 5)[:ixfParent]).to eq 16 # Set to 20% - Accent1 style
        expect(@browser.get_xf(1, 4, 5)[:ixfParent]).to eq 40 # Set to Bad style
      end

      # Names of built-in style XFs
      it 'Adds _description field for 16 default XFs' do
        expect(@browser.globals[:XF][0][:_description]).to eq :'Normal style'
        expect(@browser.globals[:XF][1][:_description]).to eq :'Row outline level 1'
        expect(@browser.globals[:XF][8][:_description]).to eq :'Column outline level 1'
        expect(@browser.globals[:XF][15][:_description]).to eq :'Default cell format'
        expect(@browser.globals[:XF][16][:_description]).to eq nil
      end

      # Decoded fStyle field
      it 'Adds _type field' do
        expect(@browser.globals[:XF][0][:_type]).to eq :stylexf
        expect(@browser.globals[:XF][15][:_type]).to eq :cellxf
        expect(@browser.get_xf(1, 5, 5)[:_type]).to eq :cellxf
      end

      # Text alignment
      it 'Decodes alc field and adds alc_d field' do
        # Can't find usage of alc == 0xFF (:ALCNIL, Alignment not specified) in cell XFs and style XFs

        f = @browser.get_xf(1, 2, 6)
        expect(f[:alc]).to eq 0
        expect(f[:alc_d]).to eq :ALCGEN

        f = @browser.get_xf(1, 3, 6)
        expect(f[:alc]).to eq 1
        expect(f[:alc_d]).to eq :ALCLEFT

        f = @browser.get_xf(1, 4, 6)
        expect(f[:alc]).to eq 2
        expect(f[:alc_d]).to eq :ALCCTR

        f = @browser.get_xf(1, 5, 6)
        expect(f[:alc]).to eq 3
        expect(f[:alc_d]).to eq :ALCRIGHT

        f = @browser.get_xf(1, 6, 6)
        expect(f[:alc]).to eq 4
        expect(f[:alc_d]).to eq :ALCFILL

        f = @browser.get_xf(1, 7, 6)
        expect(f[:alc]).to eq 5
        expect(f[:alc_d]).to eq :ALCJUST

        f = @browser.get_xf(1, 8, 6)
        expect(f[:alc]).to eq 6
        expect(f[:alc_d]).to eq :ALCCONTCTR

        f = @browser.get_xf(1, 9, 6)
        expect(f[:alc]).to eq 7
        expect(f[:alc_d]).to eq :ALCDIST
      end

      # Alignment - Text control - Wrap text
      it 'Decodes fWrap field' do
        expect(@browser.get_xf(1, 2, 7)[:fWrap]).to eq true
        expect(@browser.get_xf(1, 3, 7)[:fWrap]).to eq false
      end

      # Vertical alignment
      it 'Decodes alcV field and adds alcV_d field' do
        f = @browser.get_xf(1, 2, 8)
        expect(f[:alcV]).to eq 0
        expect(f[:alcV_d]).to eq :ALCVTOP

        f = @browser.get_xf(1, 3, 8)
        expect(f[:alcV]).to eq 1
        expect(f[:alcV_d]).to eq :ALCVCTR

        f = @browser.get_xf(1, 4, 8)
        expect(f[:alcV]).to eq 2
        expect(f[:alcV_d]).to eq :ALCVBOT

        f = @browser.get_xf(1, 5, 8)
        expect(f[:alcV]).to eq 3
        expect(f[:alcV_d]).to eq :ALCVJUST

        f = @browser.get_xf(1, 6, 8)
        expect(f[:alcV]).to eq 4
        expect(f[:alcV_d]).to eq :ALCVDIST
      end

      # Alignment - Text alignment – Justify distributed
      it 'Decodes fJustLast field' do
        expect(@browser.get_xf(1, 2, 9)[:fJustLast]).to eq true # 'Justify distributed' on
        expect(@browser.get_xf(1, 3, 9)[:fJustLast]).to eq false
      end

      # Alignment – Orientation
      it 'Decodes trot field and adds trot_d field' do
        f = @browser.get_xf(1, 2, 10) # normal
        expect(f[:trot]).to eq 0
        expect(f[:trot_d]).to eq :counterclockwise

        f = @browser.get_xf(1, 3, 10) # -45
        expect(f[:trot]).to eq 45
        expect(f[:trot_d]).to eq :counterclockwise

        f = @browser.get_xf(1, 4, 10) # +45
        expect(f[:trot]).to eq 135
        expect(f[:trot_d]).to eq :clockwise

        f = @browser.get_xf(1, 5, 10) # vertical
        expect(f[:trot]).to eq 255
        expect(f[:trot_d]).to eq :vertical
      end

      # Alignment - Text alignment – Indent
      it 'Decodes cIndent field' do
        expect(@browser.get_xf(1, 2, 11)[:cIndent]).to eq 0
        expect(@browser.get_xf(1, 3, 11)[:cIndent]).to eq 5
        expect(@browser.get_xf(1, 4, 11)[:cIndent]).to eq 15
      end

      # Alignment - Text control - Shrink to fit
      it 'Decodes fShrinkToFit field' do
        expect(@browser.get_xf(1, 2, 12)[:fShrinkToFit]).to eq true
        expect(@browser.get_xf(1, 3, 13)[:fShrinkToFit]).to eq false
      end

      # Alignment - Right-to-left - Text direction
      it 'Decodes iReadingOrder field and adds iReadingOrder_d field' do
        f = @browser.get_xf(1, 2, 13)
        expect(f[:iReadingOrder]).to eq 0
        expect(f[:iReadingOrder_d]).to eq :READING_ORDER_CONTEXT

        f = @browser.get_xf(1, 3, 13)
        expect(f[:iReadingOrder]).to eq 1
        expect(f[:iReadingOrder_d]).to eq :READING_ORDER_LTR

        f = @browser.get_xf(1, 4, 13)
        expect(f[:iReadingOrder]).to eq 2
        expect(f[:iReadingOrder_d]).to eq :READING_ORDER_RTL
      end

      # Flags specifying whether changes in parent XFs are not reflected in this XF
      # Hoping that there are records in the test file with this bit set to true:
      %i(fAtrNum fAtrFnt fAtrAlc fAtrBdr fAtrPat fAtrProt).each do |attr|
        it "Decodes #{attr} field" do
          [true, false].each do |val|
            xf = @browser.mapif(:XF, first: true) { |r, _, _, ssi| r if ssi == 0 && r[attr] == val }
            expect(xf[attr]).to eq val
          end
        end
      end

      let(:border_styles) do
        [
          { row: 2,  num_val: 0x00, decoded_val: :NONE },
          { row: 3,  num_val: 0x01, decoded_val: :THIN },
          { row: 4,  num_val: 0x02, decoded_val: :MEDIUM },
          { row: 5,  num_val: 0x03, decoded_val: :DASHED },
          { row: 6,  num_val: 0x04, decoded_val: :DOTTED },
          { row: 7,  num_val: 0x05, decoded_val: :THICK },
          { row: 8,  num_val: 0x06, decoded_val: :DOUBLE },
          { row: 9,  num_val: 0x07, decoded_val: :HAIR },
          { row: 10, num_val: 0x08, decoded_val: :MEDIUMDASHED },
          { row: 11, num_val: 0x09, decoded_val: :DASHDOT },
          { row: 12, num_val: 0x0A, decoded_val: :MEDIUMDASHDOT },
          { row: 13, num_val: 0x0B, decoded_val: :DASHDOTDOT },
          { row: 14, num_val: 0x0C, decoded_val: :MEDIUMDASHDOTDOT },
          { row: 15, num_val: 0x0D, decoded_val: :SLANTEDDASHDOTDOT },
        ]
      end

      [
        { col: 14, num_prop: :dgLeft,   decoded_prop: :dgLeft_d },
        { col: 15, num_prop: :dgRight,  decoded_prop: :dgRight_d },
        { col: 17, num_prop: :dgTop,    decoded_prop: :dgTop_d },
        { col: 18, num_prop: :dgBottom, decoded_prop: :dgBottom_d },
        { col: 26, num_prop: :dgDiag,   decoded_prop: :dgDiag_d },
      ].each do |attrs|
        it "Decodes #{attrs[:num_prop]} and adds #{attrs[:decoded_prop]} field" do
          border_styles.each do |bs|
            f = @browser.get_xf(1, bs[:row], attrs[:col])
            expect(f[attrs[:decoded_prop]]).to eq bs[:decoded_val]
            expect(f[attrs[:num_prop]]).to eq bs[:num_val]
          end
        end
      end

      # Border colors (see 2.5.161 Icv)
      let(:border_icvs) do
        {
          2 => 0x40, # default text color
          3 => 0x23, # cyan
          4 => 0x22, # yellow
        }
      end

      {
        icvLeft: 20,
        icvRight: 21,
        icvTop: 23,
        icvBottom: 24,
        icvDiag: 25,
      }.each do |field, col|
        it "Decodes #{field} field" do
          border_icvs.each do |row, icv|
            expect(@browser.get_xf(1, row, col)[field]).to eq icv
          end
        end
      end

      let(:diagonals) do
        [
          { row: 2, grbitDiag: 0, grbitDiag_d: :'No diagonal border' },
          { row: 3, grbitDiag: 1, grbitDiag_d: :'Diagonal-down border' },
          { row: 4, grbitDiag: 2, grbitDiag_d: :'Diagonal-up border' },
          { row: 5, grbitDiag: 3, grbitDiag_d: :'Both diagonal-down and diagonal-up' },
        ]
      end

      # Diagonal lines (borders)
      it 'Decodes grbitDiag and adds grbitDiag_d field' do
        diagonals.each do |attrs|
          f = @browser.get_xf(1, attrs[:row], 22)
          expect(f[:grbitDiag]).to eq(attrs[:grbitDiag])
          expect(f[:grbitDiag_d]).to eq(attrs[:grbitDiag_d])
        end
      end

      it 'Decodes fHasXFExt field' do
        expect(@browser.get_xf(1, 2, 27)[:fHasXFExt]).to eq true # Uses advanced styling (RGB font color) specified in XFExt
        expect(@browser.get_xf(1, 3, 27)[:fHasXFExt]).to eq false
      end

      let(:fill_patterns) do
        [
          { row: 2,  fls: 0x00, fls_d: :FLSNULL }, # No fill pattern
          { row: 3,  fls: 0x01, fls_d: :FLSSOLID }, # Solid
          { row: 4,  fls: 0x02, fls_d: :FLSMEDGRAY }, # 50% gray
          { row: 5,  fls: 0x03, fls_d: :FLSDKGRAY }, # 75% gray
          { row: 6,  fls: 0x04, fls_d: :FLSLTGRAY }, # 25% gray
          { row: 7,  fls: 0x05, fls_d: :FLSDKHOR }, # Horizontal stripe
          { row: 8,  fls: 0x06, fls_d: :FLSDKVER }, # Vertical stripe
          { row: 9,  fls: 0x07, fls_d: :FLSDKDOWN }, # Reverse diagonal stripe
          { row: 10, fls: 0x08, fls_d: :FLSDKUP }, # Diagonal stripe
          { row: 11, fls: 0x09, fls_d: :FLSDKGRID }, # Diagonal crosshatch
          { row: 12, fls: 0x0A, fls_d: :FLSDKTRELLIS }, # Thick diagonal crosshatch
          { row: 13, fls: 0x0B, fls_d: :FLSLTHOR }, # Thin horizontal stripe
          { row: 14, fls: 0x0C, fls_d: :FLSLTVER }, # Thin vertical stripe
          { row: 15, fls: 0x0D, fls_d: :FLSLTDOWN }, # Thin reverse diagonal stripe
          { row: 16, fls: 0x0E, fls_d: :FLSLTUP }, # Thin diagonal stripe
          { row: 17, fls: 0x0F, fls_d: :FLSLTGRID }, # Thin horizontal crosshatch
          { row: 18, fls: 0x10, fls_d: :FLSLTTRELLIS }, # Thin diagonal crosshatch
          { row: 19, fls: 0x11, fls_d: :FLSGRAY125 }, # 12.5% gray
          { row: 20, fls: 0x12, fls_d: :FLSGRAY0625 }, # 6.25% gray
        ]
      end

      # Cell pattern type
      it 'Decodes fls and adds fls_d field' do
        fill_patterns.each do |attrs|
          f = @browser.get_xf(1, attrs[:row], 28)
          expect(f[:fls]).to eq attrs[:fls]
          expect(f[:fls_d]).to eq attrs[:fls_d]
        end
      end

      # Cell pattern color (see 2.5.161 Icv)
      it 'Decodes icvFore field' do
        expect(@browser.get_xf(1, 2, 29)[:icvFore]).to eq 0x40 # default foreground color
        expect(@browser.get_xf(1, 3, 29)[:icvFore]).to eq 0x23 # cyan
      end

      # Cell background color (see 2.5.161 Icv)
      it 'Decodes icvBack field' do
        expect(@browser.get_xf(1, 2, 30)[:icvBack]).to eq 0x41 # default background color
        expect(@browser.get_xf(1, 3, 30)[:icvBack]).to eq 0x23 # cyan
      end

      # Cell is a PivotTable button
      it 'Decodes fsxButton field' do
        expect(@browser.get_xf(1, 2, 31)[:fsxButton]).to eq true
        expect(@browser.get_xf(1, 3, 31)[:fsxButton]).to eq true
        expect(@browser.get_xf(1, 4, 31)[:fsxButton]).to eq false
      end
    end

    context 'XFExt record' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'xfext.xls'))
      end

      it 'Decodes frtHeader field' do
        f = @browser.get_xfext(1, 3, 0)
        expect(f[:frtHeader][:rt]).to eq f[:_record][:id]
        expect(f[:frtHeader][:grBitFrt]).to eq({ fFrtRef: false, fFrtAlert: false })
      end

      it 'Decodes ixfe field' do
        expect(@browser.get_xfext(1, 3, 0)[:ixfe]).to be_an_instance_of Integer
        expect(@browser.get_xf(1, 3, 0)[:fHasXFExt]).to eq true
      end

      it 'Decodes cexts field' do
        expect(@browser.get_xfext(1, 3, 1)[:cexts]).to eq 2
      end

      context 'Decodes rgExt field' do
        context 'ExtProp type 0x04 (FullColorExt) for foreground (pattern) color' do
          it 'With xclrType 0 (XCLRAUTO)' # cannot find an example

          it 'With xclrType 1 (XCLRINDEXED)' # cannot find an example

          it 'With xclrType 2 (XCLRRGB)' do
            extprop = @browser.get_xfext(1, 3, 0)[:rgExt].find { |h| h[:_property] == :'cell interior foreground color' }

            expect(extprop[:extType]).to eq 4
            expect(extprop[:extType_d]).to eq :FullColorExt
            expect(extprop[:cb]).to eq 20

            propdata = extprop[:extPropData]
            expect(propdata[:xclrType]).to eq 2
            expect(propdata[:xclrType_d]).to eq :XCLRRGB
            expect(propdata[:nTintShade]).to eq 0
            expect(propdata[:nTintShade_d]).to eq 0.0
            expect(propdata[:xclrValue]).to eq :'2A8EF2FF' # rgb(42, 142, 242)
          end

          it 'With xclrType 3 (XCLRTHEMED)' do
            extprop = @browser.get_xfext(1, 4, 0)[:rgExt].find { |h| h[:_property] == :'cell interior foreground color' }

            expect(extprop[:extType]).to eq 4
            expect(extprop[:extType_d]).to eq :FullColorExt
            expect(extprop[:cb]).to eq 20

            propdata = extprop[:extPropData]
            expect(propdata[:xclrType]).to eq 3
            expect(propdata[:xclrType_d]).to eq :XCLRTHEMED # Blue, Accent 5, lighter 40%
            expect(propdata[:nTintShade]).to eq 13105
            expect(propdata[:nTintShade_d]).to eq 0.4
            expect(propdata[:xclrValue]).to eq 8
          end

          it 'With xclrType 4 (XCLRNINCHED)' # cannot find an example
        end

        context 'ExtProp type 0x05 (FullColorExt) for background color' do
          it 'With xclrType 2 (XCLRRGB)' do
            extprop = @browser.get_xfext(1, 3, 1)[:rgExt].find { |h| h[:_property] == :'cell interior background color' }
            propdata = extprop[:extPropData]
            expect(propdata[:xclrType_d]).to eq :XCLRRGB
            expect(propdata[:xclrValue]).to eq :'2A8EF2FF'
          end

          it 'With xclrType 3 (XCLRTHEMED)' do
            extprop = @browser.get_xfext(1, 4, 1)[:rgExt].find { |h| h[:_property] == :'cell interior background color' }
            propdata = extprop[:extPropData]
            expect(propdata[:xclrType_d]).to eq :XCLRTHEMED
            expect(propdata[:nTintShade_d]).to eq 0.4
            expect(propdata[:xclrValue]).to eq 8
          end
        end

        context 'ExtProp type 0x06 (XFExtGradient) for cell interior gradient' do
          def get_gradient_rgext(substream_index, row, column)
            @browser.get_xfext(substream_index, row, column)[:rgExt].find do |h|
              h[:_property] == :'cell interior gradient fill'
            end
          end

          def check_type(extprop)
            expect(extprop[:extType]).to eq 6
            expect(extprop[:extType_d]).to eq :XFExtGradient
          end

          def check_linear_grad_props(propdata)
            expect(propdata[:type]).to eq 0
            expect(propdata[:type_d]).to eq :linear
            %i(numFillToLeft numFillToRight numFillToTop numFillToBottom).each do |prop|
              expect(propdata[prop]).to eql 0.0
            end
          end

          def check_rectangular_grad_props(propdata)
            expect(propdata[:type]).to eq 1
            expect(propdata[:type_d]).to eq :rectangular
            expect(propdata[:numDegree]).to eql 0.0
          end

          def check_gradstops_2_themed(gradstops)
            (0..1).each do |i|
              expect(gradstops[i][:xclrType]).to eq 3
              expect(gradstops[i][:xclrType_d]).to eq :XCLRTHEMED
              expect(gradstops[i][:numTint]).to eql 0.0
            end

            expect(gradstops[0][:xclrValue]).to eq 0
            expect(gradstops[0][:numPosition]).to eql 0.0

            expect(gradstops[1][:xclrValue]).to eq 4
            expect(gradstops[1][:numPosition]).to eql 1.0
          end

          def check_gradstops_3_themed(gradstops)
            (0..2).each do |i|
              expect(gradstops[i][:xclrType]).to eq 3
              expect(gradstops[i][:xclrType_d]).to eq :XCLRTHEMED
              expect(gradstops[i][:numTint]).to eql 0.0
            end

            expect(gradstops[0][:xclrValue]).to eq 0
            expect(gradstops[0][:numPosition]).to eql 0.0

            expect(gradstops[1][:xclrValue]).to eq 4
            expect(gradstops[1][:numPosition]).to eql 0.5

            expect(gradstops[2][:xclrValue]).to eq 0
            expect(gradstops[2][:numPosition]).to eql 1.0
          end

          it 'With XFExtGradient type 0 (linear), 0 deg, 2 stops, themed color' do
            extprop = get_gradient_rgext(1, 3, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 96

            propdata = extprop[:extPropData]
            check_linear_grad_props(propdata)
            expect(propdata[:numDegree]).to eql 0.0
            expect(propdata[:cGradStops]).to eq 2

            check_gradstops_2_themed(propdata[:rgGradStops])
          end

          it 'With XFExtGradient type 0 (linear), 45 deg, 2 stops, themed color' do
            extprop = get_gradient_rgext(1, 4, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 96

            propdata = extprop[:extPropData]
            check_linear_grad_props(propdata)
            expect(propdata[:numDegree]).to eql 45.0
            expect(propdata[:cGradStops]).to eq 2

            check_gradstops_2_themed(propdata[:rgGradStops])
          end

          it 'With XFExtGradient type 0 (linear), 90 deg, 3 stops, themed color' do
            extprop = get_gradient_rgext(1, 5, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 118

            propdata = extprop[:extPropData]
            check_linear_grad_props(propdata)
            expect(propdata[:numDegree]).to eql 90.0
            expect(propdata[:cGradStops]).to eq 3

            check_gradstops_3_themed(propdata[:rgGradStops])
          end

          it 'With XFExtGradient type 1 (rectangular), left-top aligned, 2 stops, themed color' do
            extprop = get_gradient_rgext(1, 6, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 96

            propdata = extprop[:extPropData]
            check_rectangular_grad_props(propdata)
            %i(numFillToLeft numFillToRight numFillToTop numFillToBottom).each do |prop|
              expect(propdata[prop]).to eql 0.0
            end
            expect(propdata[:cGradStops]).to eq 2

            check_gradstops_2_themed(propdata[:rgGradStops])
          end

          it 'With XFExtGradient type 1 (rectangular), right-top aligned, 2 stops, themed color' do
            extprop = get_gradient_rgext(1, 7, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 96

            propdata = extprop[:extPropData]
            check_rectangular_grad_props(propdata)
            %i(numFillToLeft numFillToRight).each { |prop| expect(propdata[prop]).to eql 1.0 }
            %i(numFillToTop numFillToBottom).each { |prop| expect(propdata[prop]).to eql 0.0 }
            expect(propdata[:cGradStops]).to eq 2

            check_gradstops_2_themed(propdata[:rgGradStops])
          end

          it 'With XFExtGradient type 1 (rectangular), right-bottom aligned, 2 stops, themed color' do
            extprop = get_gradient_rgext(1, 8, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 96

            propdata = extprop[:extPropData]
            check_rectangular_grad_props(propdata)
            %i(numFillToLeft numFillToRight numFillToTop numFillToBottom).each do |prop|
              expect(propdata[prop]).to eql 1.0
            end
            expect(propdata[:cGradStops]).to eq 2

            check_gradstops_2_themed(propdata[:rgGradStops])
          end

          it 'With XFExtGradient type 1 (rectangular), centered, 2 stops, themed color' do
            extprop = get_gradient_rgext(1, 9, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 96

            propdata = extprop[:extPropData]
            check_rectangular_grad_props(propdata)
            %i(numFillToLeft numFillToRight numFillToTop numFillToBottom).each do |prop|
              expect(propdata[prop]).to eql 0.5
            end
            expect(propdata[:cGradStops]).to eq 2

            check_gradstops_2_themed(propdata[:rgGradStops])
          end

          it 'With XFExtGradient type 0 (linear), 0 deg, 2 stops, custom colors' do
            extprop = get_gradient_rgext(1, 10, 2)

            check_type(extprop)
            expect(extprop[:cb]).to eq 96

            propdata = extprop[:extPropData]
            check_linear_grad_props(propdata)
            expect(propdata[:numDegree]).to eql 0.0
            expect(propdata[:cGradStops]).to eq 2

            gradstops = propdata[:rgGradStops]
            expect(gradstops[0][:xclrType]).to eq 3
            expect(gradstops[0][:xclrType_d]).to eq :XCLRTHEMED
            expect(gradstops[0][:xclrValue]).to eq 8
            expect(gradstops[0][:numPosition]).to eql 0.0
            expect(gradstops[0][:numTint]).to eql 0.4

            expect(gradstops[1][:xclrType]).to eq 2
            expect(gradstops[1][:xclrType_d]).to eq :XCLRRGB
            expect(gradstops[1][:xclrValue]).to eq :'2A8EF2FF'
            expect(gradstops[1][:numPosition]).to eql 1.0
            expect(gradstops[1][:numTint]).to eql 0.0
          end
        end

        let(:border_colors) do
          [
            { row: 3, propdata: { xclrType: 2, xclrType_d: :XCLRRGB,    nTintShade_d: 0.0, xclrValue: :'2A8EF2FF' } },
            { row: 4, propdata: { xclrType: 3, xclrType_d: :XCLRTHEMED, nTintShade_d: 0.4, xclrValue: 8 } },
          ]
        end

        [
          { column: 3, type: 0x07, property: :'top cell border color' },
          { column: 4, type: 0x08, property: :'bottom cell border color' },
          { column: 5, type: 0x09, property: :'left cell border color' },
          { column: 6, type: 0x0A, property: :'right cell border color' },
          { column: 7, type: 0x0B, property: :'diagonal cell border color' },
        ].each do |attrs|
          it "ExtProp type #{attrs[:type]} (FullColorExt) for #{attrs[:property]}" do
            border_colors.each do |bc|
              extprop = @browser.get_xfext(1, bc[:row], attrs[:column])[:rgExt].find do |h|
                h[:_property] == attrs[:property]
              end

              expect(extprop[:extType]).to eq(attrs[:type])
              expect(extprop[:extType_d]).to eq(:FullColorExt)

              # nTintShade written by Excel seems to differ a bit for exactly same colors
              propdata = extprop[:extPropData]
              case extprop[:extPropData][:xclrType_d]
              when :XCLRRGB then expect(propdata[:nTintShade]).to eq 0
              else               expect(propdata[:nTintShade]).to be_within(2).of(13103) # :XCLRTHEMED
              end
              propdata.delete(:nTintShade)

              expect(propdata).to eq(bc[:propdata])
            end
          end
        end

        context 'ExtProp type 0x0D (FullColorExt) for cell text color' do
          it 'With xclrType 2 (XCLRRGB)' do
            extprop = @browser.get_xfext(1, 3, 8)[:rgExt].find { |h| h[:_property] == :'cell text color' }
            propdata = extprop[:extPropData]
            expect(propdata[:xclrType_d]).to eq :XCLRRGB
            expect(propdata[:xclrValue]).to eq :'2A8EF2FF'
          end

          it 'With xclrType 3 (XCLRTHEMED)' do
            extprop = @browser.get_xfext(1, 4, 8)[:rgExt].find { |h| h[:_property] == :'cell text color' }
            propdata = extprop[:extPropData]
            expect(propdata[:xclrType_d]).to eq :XCLRTHEMED
            expect(propdata[:nTintShade_d]).to eq 0.4
            expect(propdata[:xclrValue]).to eq 8
          end
        end

        it 'Type 0x0E (FontScheme) for font scheme' do
          xfext = @browser.globals[:XFExt][0]
          extprop = xfext[:rgExt].find { |h| h[:_property] == :'font scheme' }
          expect(extprop[:extType]).to eq 14
          expect(extprop[:extType_d]).to eq :FontScheme
          expect(extprop[:cb]).to eq 5
          expect(extprop[:extPropData][:FontScheme]).to eq 2
          expect(extprop[:extPropData][:FontScheme_d]).to eq :XFSMINOR
        end

        it 'Type 0x0F (_int1b) for text indentation level' do
          extprop = @browser.get_xfext(1, 3, 10)[:rgExt].find { |h| h[:_property] == :'text indentation level' }
          expect(extprop[:extType]).to eq 15
          expect(extprop[:extType_d]).to eq :_int1b
          expect(extprop[:cb]).to eq 6
          expect(extprop[:extPropData]).to eq 16
        end
      end
    end
  end

  context 'Workbook stream, Worksheet substream' do
    context 'Various single records' do
      before(:context) do
        @browser1 = Unxls::Biff8::Browser.new(testfile('biff8', 'empty.xls'))
        @browser2 = Unxls::Biff8::Browser.new(testfile('biff8', 'preferences.xls'))
      end

      it 'Decodes WsBool record' do
        wsb1 = @browser1.wbs[1][:WsBool] # empty.xls / Sheet1
        expect(wsb1[:fShowAutoBreaks]).to eq true # Defaults
        expect(wsb1[:fDialog]).to eq false
        expect(wsb1[:fApplyStyles]).to eq false
        expect(wsb1[:fRowSumsBelow]).to eq true
        expect(wsb1[:fColSumsRight]).to eq true
        expect(wsb1[:fFitToPage]).to eq false
        expect(wsb1[:fSyncHoriz]).to eq false
        expect(wsb1[:fSyncVert]).to eq false
        expect(wsb1[:fAltExprEval]).to eq false
        expect(wsb1[:fFormulaEntry]).to eq false

        wsb2 = @browser2.wbs[1][:WsBool] # preferences.xls / Worksheet
        expect(wsb2[:fShowAutoBreaks]).to eq false # Options - View - Window Options - Page breaks
        expect(wsb2[:fDialog]).to eq false # Sheet is a Dialog sheet
        expect(wsb2[:fApplyStyles]).to eq true # Data - Group and outline - Settings - Automatic styles
        expect(wsb2[:fRowSumsBelow]).to eq false # Data - Group and outline - Settings - Summary rows below detail
        expect(wsb2[:fColSumsRight]).to eq false # Data - Group and outline - Settings - Summary columns to right of detail
        expect(wsb2[:fFitToPage]).to eq true # Options - International - Printing - Allow A4/Letter paper resizing
        expect(wsb2[:fSyncHoriz]).to eq true # To set the attributes add and run this subroutine for the 'Worksheet' sheet:
        expect(wsb2[:fSyncVert]).to eq true # Sub Main(); Windows.Arrange xlArrangeStyleVertical, True, True, True; End Sub;
        expect(wsb2[:fAltExprEval]).to eq true # Options - Transition - Sheet options - Transition formula evaluation
        expect(wsb2[:fFormulaEntry]).to eq true # Options - Transition - Sheet options - Transition formula entry

        wsb3 = @browser2.wbs[2][:WsBool] # preferences.xls / Dialog
        expect(wsb3[:fDialog]).to eq true
      end
      
      context 'PhoneticInfo' do
        let(:record) { @browser2.wbs[1][:PhoneticInfo] }

        it 'Decodes phs field' do
          phs = record[:phs]
          font = @browser2.get_font(phs[:ifnt])
          expect(font[:fontName]).to eq 'Hiragino Kaku Gothic Pro W3'
          expect(phs[:phType]).to eq 0
          expect(phs[:phType_d]).to eq :narrow_katakana
          expect(phs[:alcH]).to eq 1
          expect(phs[:alcH_d]).to eq :left
        end

        it 'Decodes sqref field' do
          sqref = record[:sqref]
          expect(sqref[:cref]).to eq(2)

          rgrefs = sqref[:rgrefs]
          expect(rgrefs[0][:rwFirst]).to eq 2
          expect(rgrefs[0][:rwLast]).to eq 3
          expect(rgrefs[0][:colFirst]).to eq 1
          expect(rgrefs[0][:colLast]).to eq 1
          expect(rgrefs[1][:rwFirst]).to eq 5
          expect(rgrefs[1][:rwLast]).to eq 5
          expect(rgrefs[1][:colFirst]).to eq 1
          expect(rgrefs[1][:colLast]).to eq 1
        end

      end
    end

    context 'CELLTABLE-related records' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'cells.xls'))
      end

      it 'Decodes Blank record' do
        record = @browser.get_cell(1, 4, 0)
        expect(record[:rw]).to eq 4
        expect(record[:col]).to eq 0
        expect(record[:ixfe]).to be_an_instance_of(Integer)
        expect(@browser.globals[:XF][record[:ixfe]][:fls_d]).to eq :FLSGRAY125
      end

      context 'BoolErr record' do
        it 'Decodes basic fields' do
          record = @browser.get_cell(1, 25, 0)

          expect(record[:rw]).to eq 25
          expect(record[:col]).to eq 0
          expect(record[:ixfe]).to be_an_instance_of(Integer)
        end

        [
          { fError: 0, row: 25, col: 0, bBoolErr: 0x01, bBoolErr_d: :True },
          { fError: 0, row: 25, col: 1, bBoolErr: 0x00, bBoolErr_d: :False },
          { fError: 1, row: 29, col: 0, bBoolErr: 0x00, bBoolErr_d: :'#NULL!' },
          { fError: 1, row: 29, col: 1, bBoolErr: 0x07, bBoolErr_d: :'#DIV/0!' },
          { fError: 1, row: 29, col: 2, bBoolErr: 0x0F, bBoolErr_d: :'#VALUE!' },
          { fError: 1, row: 29, col: 3, bBoolErr: 0x17, bBoolErr_d: :'#REF!' },
          { fError: 1, row: 29, col: 4, bBoolErr: 0x1D, bBoolErr_d: :'#NAME?' },
          { fError: 1, row: 29, col: 5, bBoolErr: 0x24, bBoolErr_d: :'#NUM!' },
          { fError: 1, row: 29, col: 6, bBoolErr: 0x2A, bBoolErr_d: :'#N/A' },
          # { fError: 1, row: 29, col: 7, bBoolErr: 0x2B, bBoolErr_d: :'#GETTING_DATA' }, # not a permanent error
        ].each do |params|
          it "Decodes BoolErr containing #{params[:bBoolErr_d]}" do
            record = @browser.get_cell(1, params[:row], params[:col])

            expect(record[:bes][:bBoolErr]).to eq params[:bBoolErr]
            expect(record[:bes][:bBoolErr_d]).to eq params[:bBoolErr_d]
            expect(record[:bes][:fError]).to eq params[:fError]
          end
        end
      end

      context 'Formula record' do
        it 'Decodes basic fields' do
          record = @browser.get_cell(1, 33, 0)

          expect(record[:rw]).to eq 33
          expect(record[:col]).to eq 0
          expect(record[:ixfe]).to be_an_instance_of(Integer)
          expect(record[:fAlwaysCalc]).to eq false
          expect(record[:fFill]).to eq false
          expect(record[:fShrFmla]).to eq false
          expect(record[:fClearErrors]).to eq false
          expect(record[:chn]).to eq :not_implemented
          expect(record[:formula]).to eq :not_implemented
        end

        it 'Decodes Formula that returns a string' do
          record = @browser.get_cell(1, 33, 0)
          expect(record[:_value]).to eq nil
          expect(record[:_type]).to eq :string

          value = @browser.get_cell(1, 33, 0, record_type: :String) # See 2.5.133 FormulaValue
          expect(value[:string]).to eq 'Text'
        end

        it 'Decodes Formula that returns a float' do
          record = @browser.get_cell(1, 33, 1)
          expect(record[:_value]).to eq 2.0
          expect(record[:_type]).to eq :float
        end

        it 'Decodes Formula that returns a boolean' do
          record = @browser.get_cell(1, 33, 2)
          expect(record[:_value]).to eq true
          expect(record[:_type]).to eq :boolean
        end

        it 'Decodes Formula that returns an error' do
          record = @browser.get_cell(1, 33, 3)
          expect(record[:_value]).to eq :'N/A'
          expect(record[:_type]).to eq :error
        end

        it 'Decodes Formula that returns an empty string' do
          record = @browser.get_cell(1, 33, 4)
          expect(record[:_value]).to eq ''
          expect(record[:_type]).to eq :blank_string
        end
      end

      it 'Decodes LabelSst record' do
        [
          'Shared',
          'Text',
          '  Text  ',
          "Line 1\nLine 2",
          "  Line 1  \n  Line 2  "
        ].each_with_index do |text, column|
          record = @browser.get_cell(1, 21, column)
          expect(record[:rw]).to eq 21
          expect(record[:col]).to eq column
          expect(record[:ixfe]).to be_an_instance_of(Integer)

          expect(record[:isst]).to be_an_instance_of(Integer)
          sst = @browser.globals[:SST][:rgb][record[:isst]]
          expect(sst[:rgb]).to eq text
        end
      end

      it 'Decodes MulBlank record' do
        record = @browser.get_cell(1, 0, 1)

        expect(record[:rw]).to eq 0
        expect(record[:colFirst]).to eq 1
        expect(record[:colLast]).to eq 7
        expect(record[:rgixfe]).to be_an_instance_of Array
        expect(record[:rgixfe].size).to eq 7
        expect(record[:rgixfe].first).to eq 68
      end

      it 'Decodes MulRk record' do
        record = @browser.get_cell(1, 16, 1)

        expect(record[:rw]).to eq 16
        expect(record[:colFirst]).to eq 1
        expect(record[:colLast]).to eq 5
        expect(record[:rgrkrec]).to be_an_instance_of Array
        expect(record[:rgrkrec].size).to eq 5

        [1.0, -1.0, 0.01, -0.01, 0.0,].each_with_index do |rk, index|
          rkrec = record[:rgrkrec][index]
          expect(rkrec[:ixfe]).to eq 15
          expect(rkrec[:RK]).to eq rk
        end
      end

      it 'Decodes Number record' do
        [
          { col: 1, num: 1.2345 },
          { col: 2, num: -1.2345 },
        ].each do |params|
          record = @browser.get_cell(1, 7, params[:col])

          expect(record[:rw]).to eq 7
          expect(record[:col]).to eq params[:col]
          expect(record[:ixfe]).to eq 15
          expect(record[:num]).to eq params[:num]
        end
      end

      it 'Decodes RK record' do
        [
          { col: 1, RK: 1.0 },
          { col: 3, RK: -1.0 },
          { col: 5, RK: 0.01 },
          { col: 7, RK: -0.01 },
        ].each do |params|
          record = @browser.get_cell(1, 10, params[:col])

          expect(record[:rw]).to eq 10
          expect(record[:col]).to eq params[:col]
          expect(record[:ixfe]).to eq 15
          expect(record[:RK]).to eq params[:RK]
        end
      end
    end

    context 'Row, ColInfo, Dimensions, MergeCells records' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'row-colinfo-dimensions-mergecells.xls'))
      end

      context 'ColInfo record' do
        let(:default_width) { 2742 } # 10 points
        let(:default_ixfe) { 15 }
        let(:sheet_index) { 2 }
  
        it 'Decodes for column with custom width' do
          colinfo = @browser.get_colinfo(sheet_index, 1)

          expect(colinfo[:colFirst]).to eq 1
          expect(colinfo[:colLast]).to eq 1
          expect(colinfo[:coldx]).to eq 2486 # custom width (9 points)
          expect(colinfo[:ixfe]).to eq default_ixfe
          expect(colinfo[:fHidden]).to eq false
          expect(colinfo[:fUserSet]).to eq true
          expect(colinfo[:fBestFit]).to eq false
          expect(colinfo[:fPhonetic]).to eq false
          expect(colinfo[:iOutLevel]).to eq 0
          expect(colinfo[:fCollapsed]).to eq false
        end
  
        it 'Decodes for column with custom styling' do
          colinfo = @browser.get_colinfo(sheet_index, 2)
          expect(colinfo[:colFirst]).to eq 2
          expect(colinfo[:colLast]).to eq 2
          expect(colinfo[:coldx]).to eq default_width
  
          expect(colinfo[:ixfe]).to be_an_instance_of(Integer) # custom XF
          xfext = @browser.globals[:XFExt].find { |r| r[:ixfe] == colinfo[:ixfe] }
          extprop = xfext[:rgExt].find { |x| x[:_property] == :'cell interior foreground color' }
          expect(extprop[:extPropData][:xclrValue]).to eq :'2A8EF2FF'
  
          expect(colinfo[:fHidden]).to eq false
          expect(colinfo[:fUserSet]).to eq false
          expect(colinfo[:fBestFit]).to eq false
          expect(colinfo[:fPhonetic]).to eq false
          expect(colinfo[:iOutLevel]).to eq 0
          expect(colinfo[:fCollapsed]).to eq false
        end
  
        it 'Decodes for hidden column' do
          colinfo = @browser.get_colinfo(sheet_index, 3)
          expect(colinfo[:colFirst]).to eq 3
          expect(colinfo[:colLast]).to eq 3
          expect(colinfo[:coldx]).to eq 0
          expect(colinfo[:ixfe]).to eq default_ixfe
          expect(colinfo[:fHidden]).to eq true
          expect(colinfo[:fUserSet]).to eq true
          expect(colinfo[:fBestFit]).to eq false
          expect(colinfo[:fPhonetic]).to eq false
          expect(colinfo[:iOutLevel]).to eq 0
          expect(colinfo[:fCollapsed]).to eq false
        end
  
        it 'Decodes for column with AutoFit Selection' do
          colinfo = @browser.get_colinfo(sheet_index, 4)
          expect(colinfo[:colFirst]).to eq 4
          expect(colinfo[:colLast]).to eq 4
          expect(colinfo[:coldx]).to eq 3254
          expect(colinfo[:ixfe]).to eq default_ixfe
          expect(colinfo[:fHidden]).to eq false
          expect(colinfo[:fUserSet]).to eq true
          expect(colinfo[:fBestFit]).to eq true
          expect(colinfo[:fPhonetic]).to eq false
          expect(colinfo[:iOutLevel]).to eq 0
          expect(colinfo[:fCollapsed]).to eq false
        end
  
        it 'Decodes for column with phonetic fields displayed' do
          colinfo = @browser.get_colinfo(sheet_index, 5)
          expect(colinfo[:colFirst]).to eq 5
          expect(colinfo[:colLast]).to eq 5
          expect(colinfo[:coldx]).to eq default_width
          expect(colinfo[:ixfe]).to eq default_ixfe
          expect(colinfo[:fHidden]).to eq false
          expect(colinfo[:fUserSet]).to eq false
          expect(colinfo[:fBestFit]).to eq false
          expect(colinfo[:fPhonetic]).to eq true
          expect(colinfo[:iOutLevel]).to eq 0
          expect(colinfo[:fCollapsed]).to eq false
        end
  
        # Columns G and H are actually collapsed, but the flag fCollapsed is not set,
        # whereas flags fHidden and fUserSet *are* set.
        # Seems like flag fCollapsed is only set to the columns
        # that follow a set of collapsed columns.
        it 'Decodes for collapsed columns' do
          # Column G, inner outline, collapsed
          colinfo = @browser.get_colinfo(sheet_index, 6)
          expect(colinfo[:colFirst]).to eq 6
          expect(colinfo[:colLast]).to eq 6
          expect(colinfo[:coldx]).to eq default_width
          expect(colinfo[:ixfe]).to eq default_ixfe
          expect(colinfo[:fHidden]).to eq true
          expect(colinfo[:fUserSet]).to eq true
          expect(colinfo[:fBestFit]).to eq false
          expect(colinfo[:fPhonetic]).to eq false
          expect(colinfo[:iOutLevel]).to eq 2
          expect(colinfo[:fCollapsed]).to eq false
  
          # Column G, outer outline, collapsed
          colinfo = @browser.get_colinfo(sheet_index, 7)
          expect(colinfo[:colFirst]).to eq 7
          expect(colinfo[:colLast]).to eq 7
          expect(colinfo[:coldx]).to eq default_width
          expect(colinfo[:ixfe]).to eq default_ixfe
          expect(colinfo[:fHidden]).to eq true
          expect(colinfo[:fUserSet]).to eq true
          expect(colinfo[:fBestFit]).to eq false
          expect(colinfo[:fPhonetic]).to eq false
          expect(colinfo[:iOutLevel]).to eq 1
          expect(colinfo[:fCollapsed]).to eq true
  
          # Column I with default settings, directly follows collapsed column G.
          colinfo = @browser.get_colinfo(sheet_index, 8)
          expect(colinfo[:colFirst]).to eq 8
          expect(colinfo[:colLast]).to eq 8
          expect(colinfo[:coldx]).to eq default_width
          expect(colinfo[:ixfe]).to eq default_ixfe
          expect(colinfo[:fHidden]).to eq false
          expect(colinfo[:fUserSet]).to eq false
          expect(colinfo[:fBestFit]).to eq false
          expect(colinfo[:fPhonetic]).to eq false
          expect(colinfo[:iOutLevel]).to eq 0
          expect(colinfo[:fCollapsed]).to eq true
        end
      end

      context 'Row record' do
        let(:sheet_index) { 1 }

        specify 'Row record not written if row has no content' do
          expect(@browser.get_row(sheet_index, 2)).to eq nil
        end
  
        # colMic and colMac are the same for row groups of 16 records, e.g. rows 0-15, then 16-31 etc.
        # So if row 0 has cells in cols 1 and 3, and row 15 has cell in cols 5 and 10, both Row records
        # will have colMic = 1 and colMac = 10, but rows 16…31 will have their own "bounding box".
        it 'Decodes colMic and colMac fields' do
          (3..5).each do |row_index|
            row = @browser.get_row(sheet_index, row_index)
            expect(row[:colMic]).to eq 0
            expect(row[:colMac]).to eq 6
          end
        end
  
        it 'Decodes miyRw and fUnsynced fields' do
          row = @browser.get_row(sheet_index, 7)
          expect(row[:miyRw]).to eq 400
          expect(row[:fUnsynced]).to eq false
          
          row = @browser.get_row(sheet_index, 8)
          expect(row[:miyRw]).to eq 600
          expect(row[:fUnsynced]).to eq true
        end
  
        it 'Decodes iOutLevel, fCollapsed and fDyZero fields' do
          [
            { row_index: 10, iOutLevel: 4, fCollapsed: false, fDyZero: true },
            { row_index: 11, iOutLevel: 3, fCollapsed: true, fDyZero: true },
            { row_index: 12, iOutLevel: 2, fCollapsed: false, fDyZero: true },
            { row_index: 13, iOutLevel: 1, fCollapsed: true, fDyZero: false },
          ].each do |params|
            row = @browser.get_row(sheet_index, params[:row_index])
            expect(row[:iOutLevel]).to eq params[:iOutLevel]
            expect(row[:fCollapsed]).to eq params[:fCollapsed]
            expect(row[:fDyZero]).to eq params[:fDyZero]
          end
          
          expect(@browser.get_row(sheet_index, 15)[:fDyZero]).to eq true
        end
        
        it 'Decodes fGhostDirty and ixfe fields' do
          row = @browser.get_row(sheet_index, 17)
          expect(row[:fGhostDirty]).to eq true
          expect(row[:ixfe]).to be_an_instance_of Integer
          
          row_xfext = @browser.mapif(:XFExt, first: true) { |r| r if r[:ixfe] == row[:ixfe] }
          color_ext = row_xfext[:rgExt].find { |p| p[:_property] == :'cell interior foreground color' }
          expect(color_ext[:extPropData][:xclrValue]).to eq :'187CE0FF'
        end
        
        it 'Decodes fExAsc field' do
          {
            18 => false,
            19 => false,
            20 => true,
            21 => false,
            22 => true,
            23 => false,
            24 => true,
            25 => false,
            26 => true,
          }.each do |row_index, f_ex_asc|
            expect(@browser.get_row(sheet_index, row_index)[:fExAsc]).to eq f_ex_asc
          end
        end
  
        it 'Decodes fExDes' do
          {
            27 => false,
            28 => true,
            29 => true,
            30 => false,
            31 => true,
            32 => true,
            33 => false,
            34 => true,
            35 => true,
            36 => false,
            37 => true,
            38 => true,
            39 => false,
            40 => true,
            41 => true,
            42 => false,
            43 => true,
            44 => true,
            45 => false,
            46 => true,
            47 => true,
            48 => false,
          }.each do |row_index, f_ex_des|
            expect(@browser.get_row(sheet_index, row_index)[:fExDes]).to eq f_ex_des
          end
        end
  
        it 'Decodes fPhonetic field' do
          expect(@browser.get_row(sheet_index, 50)[:fPhonetic]).to eq false
          expect(@browser.get_row(sheet_index, 51)[:fPhonetic]).to eq true
        end
      end

      it 'Decodes Dimensions record' do
        record = @browser.wbs[3][:Dimensions]
        expect(record[:rwMic]).to eq 1
        expect(record[:rwMac]).to eq 8
        expect(record[:colMic]).to eq 1
        expect(record[:colMac]).to eq 6
      end
    
      it 'Decodes MergeCells record' do
        record = @browser.wbs[4][:MergeCells].first
        expect(record[:cmcs]).to eq 3
        expect(record[:rgref][0]).to eq({ rwFirst: 1, rwLast: 4, colFirst: 1, colLast: 1 })
        expect(record[:rgref][1]).to eq({ rwFirst: 1, rwLast: 1, colFirst: 3, colLast: 6 })
        expect(record[:rgref][2]).to eq({ rwFirst: 3, rwLast: 6, colFirst: 3, colLast: 6 })
      end
    end

    context 'Note, Obj, TxO records' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'note-obj-txo.xls'))
      end

      context 'Note record' do
        specify 'For cell with normal comment' do
          record = @browser.get_note(1, 1, 0)
          expect(record[:rw]).to eq 1
          expect(record[:col]).to eq 0
          expect(record[:fShow]).to eq false
          expect(record[:fRwHidden]).to eq false
          expect(record[:fColHidden]).to eq false
          expect(record[:idObj]).to be_an_instance_of Integer
          expect(record[:stAuthor]).to eq 'Microsoft Office User'
        end

        specify 'For cell with comment always shown' do
          record = @browser.get_note(1, 2, 0)
          expect(record[:rw]).to eq 2
          expect(record[:col]).to eq 0
          expect(record[:fShow]).to eq true
          expect(record[:fRwHidden]).to eq false
          expect(record[:fColHidden]).to eq false
          expect(record[:idObj]).to be_an_instance_of Integer
          expect(record[:stAuthor]).to eq 'Microsoft Office User'
        end

        specify 'For cell in hidden row' do
          record = @browser.get_note(1, 3, 0)
          expect(record[:rw]).to eq 3
          expect(record[:col]).to eq 0
          expect(record[:fShow]).to eq false
          expect(record[:fRwHidden]).to eq true
          expect(record[:fColHidden]).to eq false
          expect(record[:idObj]).to be_an_instance_of Integer
          expect(record[:stAuthor]).to eq 'Microsoft Office User'
        end

        specify 'For cell in hidden column' do
          record = @browser.get_note(1, 1, 1)
          expect(record[:rw]).to eq 1
          expect(record[:col]).to eq 1
          expect(record[:fShow]).to eq false
          expect(record[:fRwHidden]).to eq false
          expect(record[:fColHidden]).to eq true
          expect(record[:idObj]).to be_an_instance_of Integer
          expect(record[:stAuthor]).to eq 'Microsoft Office User'
        end
      end

      context 'Obj record' do
        [
          [1, 0],
          [2, 0],
          [3, 0],
          [1, 1],
        ].each do |rc|
          it 'Decodes cmo field' do
            row, column = rc
            note_obj_id = @browser.get_note(1, row, column)[:idObj]
            record = @browser.get_obj(1, note_obj_id)

            cmo = record[:cmo]
            expect(cmo[:ft]).to eq 0x15
            expect(cmo[:cb]).to eq 0x12
            expect(cmo[:ot]).to eq 0x19
            expect(cmo[:ot_d]).to eq :Note
            expect(cmo[:fLocked]).to eq true
            expect(cmo[:fPrint]).to eq true
            %i(fDefaultSize fPublished fDisabled fUIObj fRecalcObj fRecalcObjAlways).each do |field_name|
              expect(cmo[field_name]).to eq false
            end

            expect(record[:_other_fields]).to eq :not_implemented
          end
        end
      end

      context 'TxO record' do
        def get_note_txo(stream, row, col)
          note_obj_id = @browser.get_note(stream, row, col)[:idObj]
          note_obj_index = @browser.get_obj(stream, note_obj_id)[:_record][:index]
          @browser.wbs[stream][:TxO][note_obj_index]
        end

        specify 'For note with default formatting' do
          record = get_note_txo(2, 1, 0)
          expect(record[:hAlignment]).to eq 1
          expect(record[:hAlignment_d]).to eq :left
          expect(record[:vAlignment]).to eq 1
          expect(record[:vAlignment_d]).to eq :top
          expect(record[:fLockText]).to eq true
          expect(record[:fJustLast]).to eq false
          expect(record[:fSecretEdit]).to eq false
          expect(record[:rot]).to eq 0
          expect(record[:rot_d]).to eq :none
          expect(record[:ifntEmpty]).to eq 0
          expect(record[:fmla]).to eq :not_implemented
          expect(record[:text_string]).to eq 'default style, halign left, valign top, rot none'
          expect(record[:text_string].size).to eq record[:cchText]
          expect(record[:formatting_runs].size).to eq 1
          expect(record[:formatting_runs].first[:ich]).to be_an_instance_of Integer
          expect(record[:formatting_runs].first[:ifnt]).to be_an_instance_of Integer
        end

        specify 'For note with center-middle aligned, unlocked text' do
          record = get_note_txo(2, 3, 0)
          expect(record[:hAlignment]).to eq 2
          expect(record[:hAlignment_d]).to eq :centered
          expect(record[:vAlignment]).to eq 2
          expect(record[:vAlignment_d]).to eq :middle
          expect(record[:fLockText]).to eq false
          expect(record[:text_string]).to eq 'halign center, valign middle, lock text false'
        end

        specify 'For note with right-bottom aligned text' do
          record = get_note_txo(2, 5, 0)
          expect(record[:hAlignment]).to eq 3
          expect(record[:hAlignment_d]).to eq :right
          expect(record[:vAlignment]).to eq 3
          expect(record[:vAlignment_d]).to eq :bottom
          expect(record[:text_string]).to eq 'halign right, valign bottom'
        end

        specify 'For note with horizontally and vertically justified text' do
          record = get_note_txo(2, 7, 0)
          expect(record[:hAlignment]).to eq 4
          expect(record[:hAlignment_d]).to eq :justify
          expect(record[:vAlignment]).to eq 4
          expect(record[:vAlignment_d]).to eq :justify
          expect(record[:text_string]).to eq 'halign justify, valign justify'
        end

        specify 'For note with horizontally and vertically distributed text' do
          record = get_note_txo(2, 9, 0)
          expect(record[:hAlignment]).to eq 7
          expect(record[:hAlignment_d]).to eq :justify_distributed
          expect(record[:vAlignment]).to eq 7
          expect(record[:vAlignment_d]).to eq :justify_distributed
          expect(record[:text_string]).to eq 'halign distributed, valign distributed'
        end

        specify 'For note with stacked orientation' do
          record = get_note_txo(2, 1, 5)
          expect(record[:rot]).to eq 1
          expect(record[:rot_d]).to eq :stacked
          expect(record[:text_string]).to eq 'rot stacked'
        end

        specify 'For note with text rotated 90 ccw' do
          record = get_note_txo(2, 3, 5)
          expect(record[:rot]).to eq 2
          expect(record[:rot_d]).to eq :ccw_90
          expect(record[:text_string]).to eq 'rot 90 ccw'
        end

        specify 'For note with text rotated 90 cw' do
          record = get_note_txo(2, 5, 5)
          expect(record[:rot]).to eq 3
          expect(record[:rot_d]).to eq :cw_90
          expect(record[:text_string]).to eq 'rot 90 cw'
        end

        specify 'For note with no text' do
          record = get_note_txo(2, 7, 5)
          expect(@browser.get_font(record[:ifntEmpty])[:fontName]).to eq 'Tahoma'
        end

        specify 'For note with text split across several Continue records' do
          record = get_note_txo(2, 11, 0)
          expect(record[:text_string].size).to eq 16440 # characters
          expect(record[:cchText]).to eq 16445 # bytes. 5 extra bytes for 5 emojis
          expect(record[:text_string][4105..4107]).to eq 'Ж>W'
          expect(record[:text_string][12327..12329]).to eq 'W>Ж'
          expect(record[:text_string][-5..-1]).to eq 'Ж END'
        end

        specify 'For note with text with formatting runs' do
          record = get_note_txo(2, 13, 0)
          expect(record[:formatting_runs].size).to eq 5

          run = record[:formatting_runs][0]
          expect(run[:ich]).to eq 0
          expect(@browser.get_font(run[:ifnt])[:fontName]).to eq 'Candara'

          run = record[:formatting_runs][1]
          expect(run[:ich]).to eq 23
          expect(@browser.get_font(run[:ifnt])[:fontName]).to eq 'Baskerville Old Face'

          run = record[:formatting_runs][2]
          expect(run[:ich]).to eq 33
          expect(@browser.get_font(run[:ifnt])[:fontName]).to eq 'Candara'
          
          run = record[:formatting_runs][3]
          expect(run[:ich]).to eq 34
          expect(@browser.get_font(run[:ifnt])[:fontName]).to eq 'Corbel'

          run = record[:formatting_runs][4]
          expect(run[:ich]).to eq 38
          expect(@browser.get_font(run[:ifnt])[:fontName]).to eq 'Candara'
        end
      end
    end

    context 'HLink, HLinkTooltip records' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'hyperlinks.xls'))
      end

      context 'HLink record' do
        it 'Decodes basic fields' do
          record = @browser.get_hlink(1, 16, 1)

          expect(record[:rwFirst]).to eq 16
          expect(record[:rwLast]).to eq 16
          expect(record[:colFirst]).to eq 1
          expect(record[:colLast]).to eq 1
          expect(record[:hlinkClsid]).to eq :'79eac9d0-baf9-11ce-8c82-00aa004ba90b'
        end

        it 'Decodes hyperlink field for an internal link' do
          field = @browser.get_hlink(1, 16, 1)[:hyperlink]

          expect(field[:streamVersion]).to eq 2

          expect(field[:hlstmfHasMoniker]).to eq false
          expect(field[:hlstmfIsAbsolute]).to eq false
          expect(field[:hlstmfSiteGaveDisplayName]).to eq true
          expect(field[:hlstmfHasLocationStr]).to eq true
          expect(field[:hlstmfHasDisplayName]).to eq true
          expect(field[:hlstmfHasGUID]).to eq false
          expect(field[:hlstmfHasCreationTime]).to eq false
          expect(field[:hlstmfHasFrameName]).to eq false
          expect(field[:hlstmfMonikerSavedAsStr]).to eq false
          expect(field[:hlstmfAbsFromGetdataRel]).to eq false
          
          expect(field[:displayName]).to eq "#Sheet2!localname"
          expect(field[:location]).to eq "Sheet2!localname"
        end
        
        it 'Decodes hyperlink field for local file link' do
          field = @browser.get_hlink(1, 27, 1)[:hyperlink]
          
          expect(field[:hlstmfHasMoniker]).to eq true
          expect(field[:hlstmfIsAbsolute]).to eq false
          expect(field[:hlstmfSiteGaveDisplayName]).to eq true
          expect(field[:hlstmfHasLocationStr]).to eq true
          expect(field[:hlstmfHasDisplayName]).to eq true
          expect(field[:hlstmfHasGUID]).to eq false
          expect(field[:hlstmfHasCreationTime]).to eq false
          expect(field[:hlstmfHasFrameName]).to eq false
          expect(field[:hlstmfMonikerSavedAsStr]).to eq false
          expect(field[:hlstmfAbsFromGetdataRel]).to eq false
          expect(field[:displayName]).to eq 'long_unicode_filename_€.xls#A1'
          expect(field[:location]).to eq 'A1'

          ole_moniker = field[:oleMoniker]
          expect(ole_moniker[:monikerClsid]).to eq :'00000303-0000-0000-c000-000000000046'
          expect(ole_moniker[:monikerClsid_d]).to eq :FileMoniker
          
          moniker_data = ole_moniker[:data]
          expect(moniker_data[:cAnti]).to eq 0
          expect(moniker_data[:ansiLength]).to eq 28
          expect(moniker_data[:ansiPath]).to eq 'long_unicode_filename_?.xls'
          expect(moniker_data[:versionNumber]).to eq 57005
          expect(moniker_data[:cbUnicodePathSize]).to eq 60
          expect(moniker_data[:cbUnicodePathBytes]).to eq 54
          expect(moniker_data[:unicodePath]).to eq 'long_unicode_filename_€.xls'
        end

        it 'Decodes hyperlink field for a UNC path link' do
          field = @browser.get_hlink(1, 37, 1)[:hyperlink]

          expect(field[:hlstmfHasMoniker]).to eq true
          expect(field[:hlstmfIsAbsolute]).to eq true
          expect(field[:hlstmfSiteGaveDisplayName]).to eq true
          expect(field[:hlstmfHasLocationStr]).to eq false
          expect(field[:hlstmfHasDisplayName]).to eq true
          expect(field[:hlstmfHasGUID]).to eq false
          expect(field[:hlstmfHasCreationTime]).to eq false
          expect(field[:hlstmfHasFrameName]).to eq false
          expect(field[:hlstmfMonikerSavedAsStr]).to eq true
          expect(field[:hlstmfAbsFromGetdataRel]).to eq false
          expect(field[:displayName]).to eq "\\\\127.0.0.1\\PATH\\EXTP8_4.XLS"
          expect(field[:moniker]).to eq "\\\\127.0.0.1\\PATH\\EXTP8_4.XLS"
        end

        it 'Decodes hyperlink field for a HTTP link'
        it 'Decodes hyperlink field for a mail link'
      end

      context 'HLinkTooltip record'
    end
  end

  context 'Optimizations' do
    before(:context) do
      @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'row-colinfo-dimensions-mergecells.xls'))
    end

    it 'Adds cell value dimensions' do
      expect(@browser.cell_index[:dimensions][3]).to eq({ rmin: 1, rmax: 5, cmin: 1, cmax: 4 })
    end

    it 'Adds cell index for quick lookup' do
      expect(@browser.cell_index[:'3_5_4']).to eq :Formula_0
    end

    context 'Note, HLink, HLinkTooltip' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'note-obj-txo.xls'))
      end

      it 'Adds Note index for quick lookup' do
        expect(@browser.cell_index[:notes][:'1_1_0']).to eq 0
      end
    end

    context 'Note, HLink, HLinkTooltip' do
      before(:context) do
        @browser = Unxls::Biff8::Browser.new(testfile('biff8', 'hyperlinks.xls'))
      end

      it 'Adds HLink index for quick lookup' do
        expect(@browser.cell_index[:hlinks][:'1_6_1']).to eq 0
      end

      it 'Adds HLinkTooltip index for quick lookup' do
        expect(@browser.cell_index[:hlinktooltips][:'1_57_1']).to eq 0
      end
    end
  end

end
