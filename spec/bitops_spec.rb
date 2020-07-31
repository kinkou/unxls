# frozen_string_literal: true

RSpec.describe Unxls::BitOps do
  it '#set_at? detects bit flags' do
    expect(
      Unxls::BitOps.new(0b0001).set_at?(0)
    ).to eq true

    expect(
      Unxls::BitOps.new(0b0010).set_at?(1)
    ).to eq true

    expect(
      Unxls::BitOps.new(0b0100).set_at?(1)
    ).to eq false

    expect(
      Unxls::BitOps.new(0b10000000_00000000_00000000_00000000).set_at?(31)
    ).to eq true

    expect(
      Unxls::BitOps.new(0b01000000_00000000_00000000_00000000).set_at?(31)
    ).to eq false

    expect(
      Unxls::BitOps.new(0b10000000_00000000_00000000_00000000_00000000_00000000_00000000_00000000).set_at?(63)
    ).to eq true

    expect(
      Unxls::BitOps.new(0b01000000_00000000_00000000_00000000_00000000_00000000_00000000_00000000).set_at?(63)
    ).to eq false

    expect(
      Unxls::BitOps.new(0b1111).set_at?(-1)
    ).to eq nil
  end

  it '#value_at extracts values that are not byte-multiple' do
    {
      0b1100 => 2..3,
      0b11000000 => 6..7,
      0b11000000_00000000 => 14..15,
      0b11000000_00000000_00000000_00000000 => 30..31,
      0b11000000_00000000_00000000_00000000_00000000_00000000_00000000_00000000 => 62..63
    }.each do |bits, range|
      expect(
        Unxls::BitOps.new(bits).value_at(range)
      ).to eq 3
    end

    expect(
      Unxls::BitOps.new(0b0110).value_at(2..2)
    ).to eq 1

    expect(
      Unxls::BitOps.new(0b0110).value_at(0..1)
    ).to eq 2

    expect(
      Unxls::BitOps.new(0b0110).value_at(-1..0)
    ).to eq nil

    expect(
      Unxls::BitOps.new(0b0110).value_at(-2..-1)
    ).to eq nil

    expect(
      Unxls::BitOps.new(0b0110).value_at(-1..-2)
    ).to eq nil

    expect(
      Unxls::BitOps.new(0b0101).value_at(0)
    ).to eq 1

    expect(
      Unxls::BitOps.new(0b0101).value_at(1)
    ).to eq 0

    expect(
      Unxls::BitOps.new(0b1111).value_at(-1)
    ).to eq nil
  end

  it '#rol rotates left' do
    expect(
      Unxls::BitOps.new(0b1100).rol(4, 1)
    ).to eq 0b1001
  end

  it '#ror rotates right' do
    expect(
      Unxls::BitOps.new(0b1100).ror(4, 1)
    ).to eq 0b0110
  end
end
