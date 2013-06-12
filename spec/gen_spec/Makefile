SPEC_DB = 'spec.db'
DB_DUMP = 'sql/dump.sql'

.PHONY: test sdist clean

help:
	@echo
	@echo "Please use \`make <target>' where <target> is one or more of:"
	@echo "  const      - generate msoAutoShapeType constant definitions"
	@echo "  spec       - generate msoAutoShapeType spec dictionary"
	@echo "  dump_db    - generate SQL dump from which spec DB can be restored"
	@echo "  restore_db - restore spec.db from SQL dump"

const:
	@./gen_spec.py --constants msoAutoShapeType

spec:
	@./gen_spec.py msoAutoShapeType

dump_db:
	if [ -f $(DB_DUMP).bak ]; then rm -f $(DB_DUMP).bak; fi
	if [ -f $(DB_DUMP) ]; then mv $(DB_DUMP) $(DB_DUMP).bak; fi
	echo 'vacuum;' | sqlite3 $(SPEC_DB)
	echo '.dump' | sqlite3 $(SPEC_DB) > $(DB_DUMP)

restore_db:
	if [ -f $(SPEC_DB).bak ]; then rm -f $(SPEC_DB).bak; fi
	if [ -f $(SPEC_DB) ]; then mv $(SPEC_DB) $(SPEC_DB).bak; fi
	cat $(DB_DUMP) | sqlite3 $(SPEC_DB)
