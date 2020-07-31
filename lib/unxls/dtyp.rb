# frozen_string_literal: true

# [MS-DTYP]: Windows Data Types
module Unxls::Dtyp
  using Unxls::Helpers

  extend self

  # 2.3.3 FILETIME
  # @param data [String]
  # @return [Time]
  def filetime(data)
    low_byte, high_byte = data.unpack('VV')
    intervals = (high_byte << 32) | low_byte # 100-nanosecond intervals since Jan 1, 1601, UTC
    seconds = intervals / 10_000_000.0 # convert to seconds: intervals * 100.0 / 1_000_000_000
    Time.utc(1601, 1, 1) + seconds
  end

  # 2.3.4 GUID and UUID
  # A GUID, also known as a UUID, is a 16-byte structure, intended to serve as a unique identifier for an object.
  # @param data [String]
  # @return [Symbol]
  def guid(data)
    io = data.to_sio
    [
      io.read(4).reverse.unpack('H*')[0],
      io.read(2).reverse.unpack('H*')[0],
      io.read(2).reverse.unpack('H*')[0],
      io.read(8).unpack('H*')[0].insert(4, '-'),
    ].join('-').to_sym
  end

end