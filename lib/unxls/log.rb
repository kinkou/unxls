# frozen_string_literal: true

module Unxls::Log
  extend self

  # Output as a hexadecimal number (1 byte by default)
  # @return [String]
  def self.hex(num, len = 2)
    sprintf("%0#{len}X", num)
  end

  # Output as a binary number (32 bits by default)
  # @return [String]
  def self.bin(num, len = 32)
    sprintf("%0#{len}B", num)
  end

  # Format 2 byte long hex number
  # @return [String]
  def self.h2b(num)
    hex(num, 4)
  end

  # Format 4 byte long hex number
  # @return [String]
  def self.h4b(num)
    hex(num, 8)
  end

  # Format binary string
  # @return [String]
  def self.hex_str(str)
    str.bytes.map { |b| hex(b) }.join(' ')
  end

  def self.debug(data, message = nil, color = :red)
    return unless data && $DEBUG

    puts(message.send(color)) if message
    ap(data)
  end

  # @param record_params [Hash]
  def self.debug_raw_record(record_params)
    return unless record_params && $DEBUG

    params = {
      id: record_params[:id],
      size: record_params[:size],
      name: Unxls::Biff8::Record.name_by_id(record_params[:id]),
      pos: record_params[:pos],
      data_str: record_params[:data],
      data_hex: record_params[:data].map { |d| Unxls::Log.hex_str(d) },
    }

    self.debug(params, 'Reading record:')
  end

end