# frozen_string_literal: true

# [MS-OSHARED]: Office Common Data Types and Objects Structures
module Unxls::Oshared
  using Unxls::Helpers

  extend self

  # 2.3.7.4 AntiMoniker
  # This structure specifies an anti-moniker. An anti-moniker acts as the inverse of any moniker it is composed onto, effectively canceling out that moniker. In a composite moniker, anti-monikers are used to cancel out existing moniker elements, because monikers cannot be removed from a composite moniker. For more information about anti-monikers, see [MSDN-IMAMI].
  # @param io [StringIO]
  # @return [Symbol]
  def antimoniker(io)
    { count: io.read(4).unpack('V').first } # count (4 bytes): An unsigned integer that specifies the number of anti-monikers that have been composed together to create this instance. When an anti-moniker is composed with another anti-moniker, the resulting composition would have a count field equaling the sum of the two count fields of the composed anti-monikers. This value MUST be less than or equal to 1048576.
  end

  # 2.3.7.3 CompositeMoniker
  # This structure specifies a composite moniker. A composite moniker is a collection of arbitrary monikers. For more information about composite monikers see [MSDN-IMGCMI].
  # @param io [StringIO]
  # @return [Symbol]
  def compositemoniker(io)
    c_monikers = io.read(4).unpack('V').first

    {
      cMonikers: c_monikers, # cMonikers (4 bytes): An unsigned integer that specifies the count of monikers in monikerArray.
      monikerArray: c_monikers.times.map { hyperlinkmoniker(io) } # monikerArray (variable): An array of HyperlinkMonikers (section 2.3.7.2). Each array element specifies a moniker of arbitrary type.
    }
  end

  # Double-byte unterminated string used in HyperlinkMoniker
  # @param io [StringIO]
  # @param length [Integer]
  # @return [String]
  def _db_unterminated(io, length)
    Unxls::Biff8::Structure._encode_string(io.read(length))
  end

  # Double-byte 0x0000-terminated string used in HyperlinkMoniker
  # @param io [StringIO]
  # @return [String]
  def _db_zero_terminated(io)
    string = String.new
    while (char = io.read(2)) != "\x00\x00".encode(Encoding::ASCII_8BIT)
      string << char
    end
    Unxls::Biff8::Structure._encode_string(string)
  end

  # 2.3.7.8 FileMoniker
  # This structure specifies a file moniker. For more information about file monikers, see [MSDN-FM].
  # @param io [StringIO]
  # @return [Hash]
  def filemoniker(io)
    c_anti, ansi_length = io.read(6).unpack('vV')
    ansi_path = io.read(ansi_length).unpack('Z*').first.encode(Encoding::ASCII, { invalid: :replace, undef: :replace }) # 0-terminated

    result = {
      cAnti: c_anti, # cAnti (2 bytes): An unsigned integer that specifies the number of parent directory indicators at the beginning of the ansiPath field.
      ansiLength: ansi_length, # ansiLength (4 bytes): An unsigned integer that specifies the number of ANSI characters in ansiPath, including the terminating NULL character. This value MUST be less than or equal to 32767.
      ansiPath: ansi_path, # ansiPath (variable): A null-terminated array of ANSI characters that specifies the file path. The number of characters in the array is specified by ansiLength.
    }

    end_server = io.read(2).unpack('v').first
    path_is_unc = end_server != 0xFFFF
    result[:endServer] = end_server if path_is_unc # endServer (2 bytes): An unsigned integer that specifies the number of Unicode characters used to specify the server portion of the path if the path is a UNC path (including the leading "\\"). If the path is not a UNC path, this field MUST equal 0xFFFF.

    result[:versionNumber] = io.read(2).unpack('v').first # versionNumber (2 bytes): An unsigned integer that specifies the version number of this file moniker serialization implementation. MUST equal 0xDEAD.

    io.read(20) # reserved1 (16 bytes): MUST be zero and MUST be ignored; reserved2 (4 bytes): MUST be zero and MUST be ignored.

    result[:cbUnicodePathSize] = io.read(4).unpack('V').first # cbUnicodePathSize (4 bytes): An unsigned integer that specifies the size, in bytes, of cbUnicodePathBytes, usKeyValue, and unicodePath.
    return result if result[:cbUnicodePathSize] == 0 # ansiPath can be fully specified in ANSI characters

    cb_unicode_path_bytes, _ = io.read(6).unpack('Vv')
    result[:cbUnicodePathBytes] = cb_unicode_path_bytes # cbUnicodePathBytes (4 bytes): An optional unsigned integer that specifies the size, in bytes, of the unicodePath field. This field exists if and only if cbUnicodePathSize is greater than zero.
    # (skipped) usKeyValue (2 bytes): An optional unsigned integer that MUST be 3 if present. This field exists if and only if cbUnicodePathSize is greater than zero.
    result[:unicodePath] = _db_unterminated(io, cb_unicode_path_bytes) # unicodePath (variable): An optional array of Unicode characters that specifies the complete file path. This path MUST be the complete Unicode version of the file path specified in ansiPath and MUST include additional Unicode characters that cannot be completely specified in ANSI characters. The number of characters in this array is specified by cbUnicodePathBytes/2. This array MUST NOT include a terminating NULL character. This field exists if and only if cbUnicodePathSize is greater than zero.

    result
  end

  # 2.3.7.1 Hyperlink Object
  # This structure specifies a hyperlink and hyperlink-related information.
  # @param io [StringIO]
  # @return [Hash]
  def hyperlink(io)
    result = { streamVersion: io.read(4).unpack('V').first } # streamVersion (4 bytes): An unsigned integer that specifies the version number of the serialization implementation used to save this structure. This value MUST equal 2.

    attrs = Unxls::BitOps.new(io.read(4).unpack('V').first)
    result[:hlstmfHasMoniker] = attrs.set_at?(0) # A - hlstmfHasMoniker (1 bit): A bit that specifies whether this structure contains a moniker. If hlstmfMonikerSavedAsStr equals 1, this value MUST equal 1.
    result[:hlstmfIsAbsolute] = attrs.set_at?(1) # B - hlstmfIsAbsolute (1 bit): A bit that specifies whether this hyperlink is an absolute path or relative path.
    result[:hlstmfSiteGaveDisplayName] = attrs.set_at?(2) # C - hlstmfSiteGaveDisplayName (1 bit): A bit that specifies whether the creator of the hyperlink specified a display name.
    result[:hlstmfHasLocationStr] = attrs.set_at?(3) # D - hlstmfHasLocationStr (1 bit): A bit that specifies whether this structure contains a hyperlink location.
    result[:hlstmfHasDisplayName] = attrs.set_at?(4) # E - hlstmfHasDisplayName (1 bit): A bit that specifies whether this structure contains a display name.
    result[:hlstmfHasGUID] = attrs.set_at?(5) # F - hlstmfHasGUID (1 bit): A bit that specifies whether this structure contains a GUID as specified by [MS-DTYP].
    result[:hlstmfHasCreationTime] = attrs.set_at?(6) # G - hlstmfHasCreationTime (1 bit): A bit that specifies whether this structure contains the creation time of the file that contains the hyperlink.
    result[:hlstmfHasFrameName] = attrs.set_at?(7) # H - hlstmfHasFrameName (1 bit): A bit that specifies whether this structure contains a target frame name.
    result[:hlstmfMonikerSavedAsStr] = attrs.set_at?(8) # I - hlstmfMonikerSavedAsStr (1 bit): A bit that specifies whether the moniker was saved as a string.
    result[:hlstmfAbsFromGetdataRel] = attrs.set_at?(9) # J - hlstmfAbsFromGetdataRel (1 bit): A bit that specifies whether the hyperlink specified by this structure is an absolute path generated from a relative path.
    # reserved (22 bits): MUST be zero and MUST be ignored.

    result[:displayName] = hyperlinkstring(io) if result[:hlstmfHasDisplayName] # displayName (variable): An optional HyperlinkString (section 2.3.7.9) that specifies the display name for the hyperlink. MUST exist if and only if hlstmfHasDisplayName equals 1.
    result[:targetFrameName] = hyperlinkstring(io) if result[:hlstmfHasFrameName] # targetFrameName (variable): An optional HyperlinkString (section 2.3.7.9) that specifies the target frame. MUST exist if and only if hlstmfHasFrameName equals 1.
    result[:moniker] = hyperlinkstring(io) if result[:hlstmfHasMoniker] && result[:hlstmfMonikerSavedAsStr] # moniker (variable): An optional HyperlinkString (section 2.3.7.9) that specifies the hyperlink moniker. MUST exist if and only if hlstmfHasMoniker equals 1 and hlstmfMonikerSavedAsStr equals 1.

    if result[:hlstmfHasMoniker] && !result[:hlstmfMonikerSavedAsStr]
      result[:oleMoniker] = hyperlinkmoniker(io) # oleMoniker (variable): An optional HyperlinkMoniker (section 2.3.7.2) that specifies the hyperlink moniker. MUST exist if and only if hlstmfHasMoniker equals 1 and hlstmfMonikerSavedAsStr equals 0.
    end

    result[:location] = hyperlinkstring(io) if result[:hlstmfHasLocationStr] # location (variable): An optional HyperlinkString (section 2.3.7.9) that specifies the hyperlink location. MUST exist if and only if hlstmfHasLocationStr equals 1.
    result[:guid] = io.read(16) if result[:hlstmfHasGUID] # guid (16 bytes): An optional GUID (see 2.3.4 GUID and UUID) as specified by [MS-DTYP] that identifies this hyperlink. MUST exist if and only if hlstmfHasGUID equals 1.
    result[:fileTime] = Unxls::Dtyp.filetime(io.read(8)) if result[:hlstmfHasCreationTime] # fileTime (8 bytes): An optional FileTime structure as specified by [MS-DTYP] that specifies the UTC file creation time. MUST exist if and only if hlstmfHasCreationTime equals 1.

    result
  end

  # 2.3.7.2 HyperlinkMoniker
  # This structure specifies a hyperlink moniker.
  # @param io [StringIO]
  # @return [Hash]
  def hyperlinkmoniker(io)
    id = io.read(4).unpack('V').first
    io.pos -= 4

    moniker_type = {
      0x79eac9e0 => :URLMoniker,
      0x00000303 => :FileMoniker,
      0x00000309 => :CompositeMoniker,
      0x00000305 => :AntiMoniker,
      0x00000304 => :ItemMoniker,
    }[id]

    result = {
      monikerClsid: Unxls::Dtyp.guid(io.read(16)), # monikerClsid (16 bytes): A class identifier (CLSID) that specifies the Component Object Model (COM) component that saved this structure.
      monikerClsid_d: moniker_type,
    }
    
    processor_method = moniker_type.downcase
    result[:data] = self.send(processor_method, io) if self.respond_to?(processor_method) # data (variable): A moniker of the type specified by monikerClsid.

    result
  end

  # 2.3.7.9 HyperlinkString
  # This structure specifies a string for a hyperlink.
  # @param io [StringIO]
  # @return [String]
  def hyperlinkstring(io)
    io.read(4) # length (4 bytes): An unsigned integer that specifies the number of Unicode characters in the string field, including the null-terminating character.
    _db_zero_terminated(io) # string (variable): A null-terminated array of Unicode characters. The number of characters in the array is specified by the length field.
  end

  # 2.3.7.5 ItemMoniker
  # This structure specifies an item moniker. Item monikers are used to identify objects within containers, such as a portion of a document, an embedded object within a compound document, or a range of cells within a spreadsheet. For more information about item monikers, see [MSDN-IMCOM].
  # @param io [StringIO]
  # @return [Symbol]
  def itemmoniker(io)
    :not_implemented
  end

  # 2.3.7.7 URICreateFlags
  # This structure specifies creation flags for an [RFC3986] compliant URI. For more information about URI creation flags, see [MSDN-CreateUri].
  # @param data [String]
  # @return [Hash]
  def uricreateflags(data)
    attrs = Unxls::BitOps.new(data.unpack('V').first)

    {
      createAllowRelative: attrs.set_at?(0), # A - createAllowRelative (1 bit): A bit that specifies that if the URI scheme is unspecified and not implicitly "file," a relative scheme is assumed during creation of the URI.
      createAllowImplicitWildcardScheme: attrs.set_at?(1), # B - createAllowImplicitWildcardScheme (1 bit): A bit that specifies that if the URI scheme is unspecified and not implicitly "file," a wildcard scheme is assumed during creation of the URI.
      createAllowImplicitFileScheme: attrs.set_at?(2), # C - createAllowImplicitFileScheme (1 bit): A bit that specifies that if the URI scheme is unspecified and the URI begins with a drive letter or a UNC path, a file scheme is assumed during creation of the URI.
      createNoFrag: attrs.set_at?(3), # D - createNoFrag (1 bit): A bit that specifies that if a URI query string is present, the URI fragment is not looked for during creation of the URI.
      createNoCanonicalize: attrs.set_at?(4), # E - createNoCanonicalize (1 bit): A bit that specifies that the scheme, host, authority, path, and fragment will not be canonicalized during creation of the URI. This value MUST be 0 if createCanonicalize equals 1.
      createCanonicalize: attrs.set_at?(5), # F - createCanonicalize (1 bit): A bit that specifies that the scheme, host, authority, path, and fragment will be canonicalized during creation of the URI. This value MUST be 0 if createNoCanonicalize equals 1.
      createFileUseDosPath: attrs.set_at?(6), # G - createFileUseDosPath (1 bit): A bit that specifies that MS-DOS path compatibility mode will be used during creation of file URIs.
      createDecodeExtraInfo: attrs.set_at?(7), # H - createDecodeExtraInfo (1 bit): A bit that specifies that percent encoding and percent decoding canonicalizations will be performed on the URI query and URI fragment during creation of the URI. This field takes precedence over the createNoCanonicalize field.
      createNoDecodeExtraInfo: attrs.set_at?(8), # I - createNoDecodeExtraInfo (1 bit): A bit that specifies that percent encoding and percent decoding canonicalizations will not be performed on the URI query and URI fragment during creation of the URI. This field takes precedence over the createCanonicalize field.
      createCrackUnknownSchemes: attrs.set_at?(9), # J - createCrackUnknownSchemes (1 bit): A bit that specifies that hierarchical URIs with unrecognized URI schemes will be treated like hierarchical URIs during creation of the URI. This value MUST be 0 if createNoCrackUnknownSchemes equals 1.
      createNoCrackUnknownSchemes: attrs.set_at?(10), # K - createNoCrackUnknownSchemes (1 bit): A bit that specifies that hierarchical URIs with unrecognized URI schemes will be treated like opaque URIs during creation of the URI. This value MUST be 0 if createCrackUnknownSchemes equals 1.
      createPreProcessHtmlUri: attrs.set_at?(11), # L - createPreProcessHtmlUri (1 bit): A bit that specifies that preprocessing will be performed on the URI to remove control characters and white space during creation of the URI. This value MUST be 0 if createNoPreProcessHtmlUri equals 1.
      createNoPreProcessHtmlUri: attrs.set_at?(12), # M - createNoPreProcessHtmlUri (1 bit): A bit that specifies that preprocessing will not be performed on the URI to remove control characters and white space during creation of the URI. This value MUST be 0 if createPreProcessHtmlUri equals 1.
      createIESettings: attrs.set_at?(13), # N - createIESettings (1 bit): A bit that specifies that registry settings will be used to determine default URL parsing behavior during creation of the URI. This value MUST be 0 if createNoIESettings equals 1.
      createNoIESettings: attrs.set_at?(14), # O - createNoIESettings (1 bit): A bit that specifies that registry settings will not be used to determine default URL parsing behavior during creation of the URI. This value MUST be 0 if createIESettings equals 1.
      createNoEncodeForbiddenCharacters: attrs.set_at?(15), # P - createNoEncodeForbiddenCharacters (1 bit): A bit that specifies that URI characters forbidden in [RFC3986] will not be percent-encoded during creation of the URI.
      # reserved (16 bits): MUST be zero and MUST be ignored.
    }
  end

  # 2.3.7.6 URLMoniker
  # This structure specifies a URL moniker. For more information about URL monikers, see [MSDN-URLM].
  # @param io [StringIO]
  # @return [Hash]
  def urlmoniker(io)
    length = io.read(4).unpack('V').first
    moniker_data_io = io.read(length).to_sio

    result = {
      length: length, # length (4 bytes): An unsigned integer that specifies the size of this structure in bytes, excluding the size of the length field.
      url: _db_zero_terminated(moniker_data_io) # url (variable): A null-terminated array of Unicode characters that specifies the URL. The number of characters in the array is determined by the position of the terminating NULL character.
    }
    return result if moniker_data_io.eof?

    result[:serialGUID] = Unxls::Dtyp.guid(moniker_data_io.read(16)) # serialGUID (16 bytes): An optional GUID as specified by [MS-DTYP] for this implementation of the URL moniker serialization. This field MUST equal {0xF4815879, 0x1D3B, 0x487F, 0xAF, 0x2C, 0x82, 0x5D, 0xC4, 0x85, 0x27, 0x63} if present.
    result[:serialVersion] = moniker_data_io.read(4).unpack('V').first # serialVersion (4 bytes): An optional unsigned integer that specifies the version number of this implementation of the URL moniker serialization. This field MUST equal 0 if present.
    result[:uriFlags] = uricreateflags(moniker_data_io.read(4)) # uriFlags (4 bytes): An optional URICreateFlags structure (section 2.3.7.7) that specifies creation flags for an [RFC3986] compliant URI.

    result
  end

end