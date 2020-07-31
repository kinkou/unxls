#include <ruby.h>
#include "extconf.h"

static unsigned long
make_mask_internal(int length, int offset) {
  unsigned long mask = 0;

  while (length > 0) {
    mask <<= 1;
    mask |= 1;
    length--;
  }

  return mask << offset;
}

VALUE
rb_unxls_bit_ops_make_mask(VALUE self, VALUE length, VALUE offset) {
  int length_c = NUM2INT(length);
  int offset_c = NUM2INT(offset);
  return ULONG2NUM(make_mask_internal(length_c, offset_c));
}

VALUE
rb_unxls_bit_ops_set_at(VALUE self, VALUE index) {
  VALUE bits = rb_iv_get(self, "@bits");
  unsigned long bits_c = NUM2ULONG(bits);

  int index_c = NUM2INT(index);

  unsigned long mask = 1;

  if (index_c < 0) {
    return Qnil;
  }

  if ((bits_c & (mask << index_c)) != 0) {
    return Qtrue;
  } else {
    return Qfalse;
  }
}

VALUE
rb_unxls_bit_ops_value_at(VALUE self, VALUE range) {
  VALUE bits = rb_iv_get(self, "@bits");
  unsigned long bits_c = NUM2ULONG(bits);

  VALUE length;
  int length_c;

  VALUE offset;
  int offset_c;

  unsigned long mask;

  if (FIXNUM_P(range)) {
    length_c = 1;
    offset_c = NUM2INT(range);
  } else {
    length = rb_funcall(range, rb_intern("size"), 0, Qnil);
    length_c = NUM2INT(length);
    offset = rb_funcall(range, rb_intern("min"), 0, Qnil);
    if (NIL_P(offset)) {
      return Qnil;
    }
    offset_c = NUM2INT(offset);
  }

  if (length_c < 1 || offset_c < 0) {
    return Qnil;
  }

  mask = make_mask_internal(length_c, offset_c);

  bits_c = (bits_c & mask) >> offset_c;

  return ULONG2NUM(bits_c);
}

void
Init_bit_ops() {
  VALUE module = rb_define_module("Unxls");
  VALUE klass = rb_define_class_under(module, "BitOps", rb_cObject);

  rb_define_method(klass, "set_at?", rb_unxls_bit_ops_set_at, 1);
  rb_define_method(klass, "value_at", rb_unxls_bit_ops_value_at, 1);
  rb_define_private_method(klass, "make_mask", rb_unxls_bit_ops_make_mask, 2);
}
