extends ConfigTable

class DataType:
    extends Reference
	var test_name_1
	var test_name_2
	var test_name_4

    func _init(field_value_map := {}):
        for key in field_value_map.keys():
            _set(key, field_value_map[key])

func _get_data_table():
    # DataType.new({})
    return [
			DataType.new({"test_name_1": "123",}),
        ]

func by(field_name, v) -> DataType:
    return ._by(field_name, v) as DataType

func _get_data_head_def():
    return [
		"test_name_1",
		"test_name_2",
		"test_name_4",
    ]

# func by_field1(v) -> DataType:
#   return by("field1", v)
func by_test_name_1(v) -> DataType:
	return by("test_name_1", v)

func by_test_name_2(v) -> DataType:
	return by("test_name_2", v)

func by_test_name_4(v) -> DataType:
	return by("test_name_4", v)

