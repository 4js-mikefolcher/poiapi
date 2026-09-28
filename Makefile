# Makefile for poiapi — generated from poiapi.4pw
#
# Builds:
#   - lib group  : lib/*.4gl  -> com/fourjs/poiapi/*.42m  (package modules, no link)
#                  lib/package.xml -> com/fourjs/poiapi/package.xml (XML copy rule)
#   - app group  : src/fgl_excel_api_test.4gl -> bin/fgl_excel_api_test.42m + .42r
#                  src/*.per -> bin/*.42f

# ---------------------------------------------------------------------------
# Host differences
#
# Two separate questions, which are easy to confuse:
#
#   * The list separator for CLASSPATH and FGLLDPATH follows the PLATFORM,
#     because it is the Java and Genero runtimes that read those variables:
#     ";" on Windows, ":" everywhere else.
#
#   * The file commands follow the SHELL. GNU make drives cmd.exe on Windows
#     unless a Unix shell is installed, and cmd.exe has neither "mkdir -p"
#     nor "cp" and wants backslashes in paths. Under MSYS or Git Bash,
#     though, the Unix commands work perfectly well on Windows, so testing
#     the OS here would break those. The probe below tests the shell
#     instead: cmd.exe echoes back the quotes that a Unix shell strips.
# ---------------------------------------------------------------------------
empty  :=
space  := $(empty) $(empty)
bslash := $(empty)\$(empty)

ifeq ($(OS),Windows_NT)
  PATHSEP := ;
else
  PATHSEP := :
endif

ifeq ($(shell echo "probe"),"probe")
  native  = $(subst /,$(bslash),$1)
  MKDIR_P = if not exist "$(call native,$1)" mkdir "$(call native,$1)"
  COPY    = copy /Y "$(call native,$1)" "$(call native,$2)" >nul
  RM_F    = -del /q $(call native,$1) 2>nul
else
  native  = $1
  MKDIR_P = mkdir -p $1
  COPY    = cp $1 $2
  RM_F    = rm -f $1
endif

# ---------------------------------------------------------------------------
# Environment (mirrors the 4pw Application/Library environment settings)
# ---------------------------------------------------------------------------
# Absolute jar paths (like $(ProjectDir) in the 4pw) so recipes that cd
# into bin/ — fgllink, make run — still resolve them
JARDIR   := $(CURDIR)/.fglpkg/jars
JARS     := $(wildcard $(JARDIR)/*.jar)

# 4pw: CLASSPATH=<jars>;$(CLASSPATH) — jars first, inherited value appended
export CLASSPATH  := $(subst $(space),$(PATHSEP),$(strip $(JARS)))$(if $(CLASSPATH),$(PATHSEP)$(CLASSPATH))
# 4pw: FGLLDPATH=$(FGLLDPATH);$(ProjectDir) — project root appended,
# so IMPORT FGL com.fourjs.poiapi.* resolves
export FGLLDPATH  := $(if $(FGLLDPATH),$(FGLLDPATH)$(PATHSEP))$(CURDIR)

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
	$(call COPY,$<,$@)

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
	$(call MKDIR_P,$@)

# --- run (default configuration passes "web" as command line argument) -------
run: all
	cd $(BINDIR) && fglrun $(APP).42r web

# --- clean --------------------------------------------------------------------
clean:
	$(call RM_F,$(LIB42M) $(PKGXML))
	$(call RM_F,$(APP42M) $(APP42R) $(FORMS42F))
