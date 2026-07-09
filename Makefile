# Makefile for poiapi — generated from poiapi.4pw
#
# Builds:
#   - lib group  : lib/*.4gl  -> com/fourjs/poiapi/*.42m  (package modules, no link)
#                  lib/package.xml -> com/fourjs/poiapi/package.xml (XML copy rule)
#   - app group  : src/fgl_excel_api_test.4gl -> bin/fgl_excel_api_test.42m + .42r
#                  src/*.per -> bin/*.42f

# ---------------------------------------------------------------------------
# Environment (mirrors the 4pw Application/Library environment settings)
# ---------------------------------------------------------------------------
# Absolute jar paths (like $(ProjectDir) in the 4pw) so recipes that cd
# into bin/ — fgllink, make run — still resolve them
JARDIR   := $(CURDIR)/.fglpkg/jars
JARS     := $(wildcard $(JARDIR)/*.jar)

empty :=
space := $(empty) $(empty)

# 4pw: CLASSPATH=<jars>;$(CLASSPATH) — jars first, inherited value appended
# (Unix ':' separator; the 4pw uses ';' for Studio dirlists)
export CLASSPATH  := $(subst $(space),:,$(strip $(JARS)))$(if $(CLASSPATH),:$(CLASSPATH))
# 4pw: FGLLDPATH=$(FGLLDPATH);$(ProjectDir) — project root appended,
# so IMPORT FGL com.fourjs.poiapi.* resolves
export FGLLDPATH  := $(if $(FGLLDPATH),$(FGLLDPATH):)$(CURDIR)

FGLCOMP  := fglcomp -M
FGLFORM  := fglform -M
FGLLINK  := fgllink

# ---------------------------------------------------------------------------
# Files
# ---------------------------------------------------------------------------
PKGDIR   := com/fourjs/poiapi
BINDIR   := bin

LIBMODS  := fgl_excel \
            fgl_structures \
            fgl_spreadsheet_helper \
            fgl_spreadsheet_api \
            fgl_spreadsheet_interface \
            fgl_spreadsheet_xapi \
            fgl_table_export

LIB42M   := $(addprefix $(PKGDIR)/,$(addsuffix .42m,$(LIBMODS)))
PKGXML   := $(PKGDIR)/package.xml

FORMS    := fgl_excel_form fgl_excel_form_xtend fgl_excel_menu_table
FORMS42F := $(addprefix $(BINDIR)/,$(addsuffix .42f,$(FORMS)))

APP      := fgl_excel_api_test
APP42M   := $(BINDIR)/$(APP).42m
APP42R   := $(BINDIR)/$(APP).42r

# ---------------------------------------------------------------------------
# Targets
# ---------------------------------------------------------------------------
.PHONY: all lib app forms run clean

all: lib app forms

lib: $(LIB42M) $(PKGXML)

app: $(APP42R)

forms: $(FORMS42F)

# --- lib modules: sources declare PACKAGE com.fourjs.poiapi, so fglcomp
# appends the package path to the output dir — output base is the project root
$(PKGDIR)/%.42m: lib/%.4gl | $(PKGDIR)
	$(FGLCOMP) -o . $<

# Inter-module dependencies (IMPORT FGL com.fourjs.poiapi.*)
$(PKGDIR)/fgl_spreadsheet_helper.42m:    $(PKGDIR)/fgl_excel.42m
$(PKGDIR)/fgl_spreadsheet_api.42m:       $(PKGDIR)/fgl_excel.42m \
                                         $(PKGDIR)/fgl_spreadsheet_helper.42m
$(PKGDIR)/fgl_spreadsheet_interface.42m: $(PKGDIR)/fgl_spreadsheet_helper.42m
$(PKGDIR)/fgl_spreadsheet_xapi.42m:      $(PKGDIR)/fgl_excel.42m \
                                         $(PKGDIR)/fgl_spreadsheet_helper.42m \
                                         $(PKGDIR)/fgl_spreadsheet_api.42m \
                                         $(PKGDIR)/fgl_structures.42m
$(PKGDIR)/fgl_table_export.42m:          $(PKGDIR)/fgl_spreadsheet_helper.42m \
                                         $(PKGDIR)/fgl_spreadsheet_xapi.42m

# XML copy build rule from the 4pw
$(PKGXML): lib/package.xml | $(PKGDIR)
	cp $< $@

# --- application ------------------------------------------------------------
$(APP42M): src/$(APP).4gl $(LIB42M) | $(BINDIR)
	$(FGLCOMP) -o $(BINDIR) $<

$(APP42R): $(APP42M)
	cd $(BINDIR) && $(FGLLINK) -o $(APP).42r $(APP).42m

# --- forms ------------------------------------------------------------------
$(BINDIR)/%.42f: src/%.per | $(BINDIR)
	$(FGLFORM) -o $(BINDIR) $<

# --- directories --------------------------------------------------------------
$(PKGDIR) $(BINDIR):
	mkdir -p $@

# --- run (default configuration passes "web" as command line argument) -------
run: all
	cd $(BINDIR) && fglrun $(APP).42r web

# --- clean --------------------------------------------------------------------
clean:
	rm -f $(LIB42M) $(PKGXML)
	rm -f $(APP42M) $(APP42R) $(FORMS42F)
